import { IFitmentReport } from './models';
import { ANTHROPIC_API_KEY } from './AIConfig';

const MODEL = 'claude-haiku-4-5-20251001';
const MAX_TOKENS = 1024;
const API_URL = 'https://api.anthropic.com/v1/messages';

const PDF_DATA_URI_PREFIX = 'data:application/pdf;base64,';

const SYSTEM_PROMPT = `You are an expert recruitment AI assistant. You will be given a candidate's \
resume — either as a PDF document or as plain text — and a set of job requirements. Read the ENTIRE \
resume, including every page of multi-page PDFs, before forming your assessment.

Respond with ONLY valid JSON: no explanation, no markdown formatting, no code fences, no text before \
or after the JSON object. Every field in the requested schema must be present and have the correct \
type — use "Unknown", empty arrays, or 0 for anything you cannot determine rather than omitting a \
field or breaking the JSON structure.

If the resume has no extractable text (e.g. it is a scanned image with no text layer, is corrupted, \
or is blank), still return the full JSON structure: set "fitmentScore" to 0, "matchingSkills" and \
"missingSkills" to empty arrays, "recommendation" to "Not Recommended", "experienceMatch" to "below", \
and use "summary" to state plainly that the resume content could not be read and HR should request a \
text-based copy. Never refuse to answer or return anything other than the JSON object.

Be precise and objective about skill matching.`;

type IContentPart =
  | { type: 'text'; text: string }
  | { type: 'document'; source: { type: 'base64'; media_type: string; data: string } };

interface IAnthropicMessage {
  role: 'user' | 'assistant';
  content: string | IContentPart[];
}

interface IAnthropicRequest {
  model: string;
  max_tokens: number;
  temperature?: number;
  system: string;
  messages: IAnthropicMessage[];
}

interface IAnthropicResponse {
  content: Array<{ type: string; text: string }>;
  usage: { input_tokens: number; output_tokens: number };
}

async function callClaude(request: IAnthropicRequest): Promise<string> {
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
      'content-type': 'application/json',
    },
    body: JSON.stringify(request),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Claude API error ${response.status}: ${errorText}`);
  }

  const data = (await response.json()) as IAnthropicResponse;

  console.log('[AIService] Token usage:', {
    input: data.usage.input_tokens,
    output: data.usage.output_tokens,
    model: request.model,
  });

  const content = data.content[0];
  if (!content || content.type !== 'text') {
    throw new Error('Unexpected response type from Claude API');
  }
  return content.text;
}

function buildScreeningPrompt(
  jobTitle: string,
  department: string,
  requiredSkills: string,
  experience: string
): string {
  return `Job Title: ${jobTitle}
Department: ${department}
Required Skills: ${requiredSkills}
Experience Required: ${experience}

Analyse the resume above against these requirements and return exactly this JSON structure with no other text \
(no markdown, no code fences, nothing before or after the braces):
{
  "candidateName": "full name extracted from resume, or \\"Unknown\\" if not present",
  "contactEmail": "email address from resume, or \\"\\" if not present",
  "contactPhone": "phone number from resume, or \\"\\" if not present",
  "matchingSkills": ["skill1", "skill2"],
  "missingSkills": ["skill1", "skill2"],
  "fitmentScore": 75,
  "experienceMatch": "meets",
  "summary": "2-3 sentence objective assessment of this candidate",
  "recommendation": "Recommended"
}

Field constraints — follow these exactly:
- "fitmentScore" is an integer from 0 to 100.
- "experienceMatch" must be exactly one of: "meets", "exceeds", "below" — no other values.
- "recommendation" must be exactly one of: "Recommended", "Maybe", "Not Recommended" — no other values.
- "matchingSkills" / "missingSkills" are arrays of short strings (skill names only), never null.`;
}

function buildUserContent(
  resume: string,
  jobTitle: string,
  department: string,
  requiredSkills: string,
  experience: string
): string | IContentPart[] {
  const prompt = buildScreeningPrompt(jobTitle, department, requiredSkills, experience);

  if (resume.startsWith(PDF_DATA_URI_PREFIX)) {
    // Resume is a raw PDF — send as a document so Claude can read it natively
    const base64Data = resume.slice(PDF_DATA_URI_PREFIX.length);
    return [
      {
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: base64Data },
      },
      { type: 'text', text: prompt },
    ];
  }

  // Plain text resume — embed inline as before
  return `Resume:\n${resume}\n\n${prompt}`;
}

function parseReport(raw: string): IFitmentReport {
  const cleaned = raw
    .replace(/```json/gi, '')
    .replace(/```/g, '')
    .trim();

  const parsed = JSON.parse(cleaned) as Record<string, unknown>;

  return {
    candidateName: String(parsed.candidateName ?? ''),
    contactEmail: String(parsed.contactEmail ?? ''),
    contactPhone: String(parsed.contactPhone ?? ''),
    matchingSkills: Array.isArray(parsed.matchingSkills)
      ? (parsed.matchingSkills as unknown[]).map(String)
      : [],
    missingSkills: Array.isArray(parsed.missingSkills)
      ? (parsed.missingSkills as unknown[]).map(String)
      : [],
    fitmentScore: Number(parsed.fitmentScore ?? 0),
    experienceMatch: (['meets', 'exceeds', 'below'].indexOf(String(parsed.experienceMatch)) !== -1
      ? parsed.experienceMatch
      : 'meets') as IFitmentReport['experienceMatch'],
    summary: String(parsed.summary ?? ''),
    recommendation: (['Recommended', 'Maybe', 'Not Recommended'].indexOf(String(parsed.recommendation)) !== -1
      ? parsed.recommendation
      : 'Maybe') as IFitmentReport['recommendation'],
  };
}

export class AIService {
  public async generateJobDescription(
    templateText: string,
    job: {
      jobTitle: string;
      department: string;
      jobLocation: string;
      jobType: string;
      mustHaveSkills: string;
      goodToHaveSkills: string;
      experience: string;
      applyUrl?: string;
    }
  ): Promise<string> {
    const templateSection = templateText.trim()
      ? `Use the following document as your format and style guide — mirror its structure, tone, and section headings:\n\n${templateText.slice(0, 6000)}\n\n---\n\n`
      : '';

    const applySection = job.applyUrl
      ? `\n\nAt the very end of the job description, add a "HOW TO APPLY" section with this exact apply link: ${job.applyUrl}`
      : '';

    const userPrompt = `${templateSection}Generate a complete, professional job description for this role:

Job Title: ${job.jobTitle}
Department: ${job.department}
Location: ${job.jobLocation}
Job Type: ${job.jobType}
Experience Required: ${job.experience}
Must Have Skills: ${job.mustHaveSkills}
Good To Have Skills: ${job.goodToHaveSkills || 'None specified'}

${templateText.trim()
  ? 'Follow the template structure above exactly. Replace placeholder content with real, specific content for this role.'
  : `Include these sections:
1. About the Role
2. Key Responsibilities (6–8 bullet points)
3. Must Have Requirements
4. Good To Have Skills
5. What We Offer
6. How to Apply

Write in a professional, engaging, inclusive tone. Be specific and concrete.`}${applySection}

Return plain text only — no markdown, no asterisks, no special formatting.`;

    return callClaude({
      model: MODEL,
      max_tokens: 2048,
      temperature: 0.7,
      system: 'You are an expert HR professional and technical writer. Create compelling, clear, inclusive job descriptions. Write in plain text with clear section headers using ALL CAPS or numbered headings.',
      messages: [{ role: 'user', content: userPrompt }],
    });
  }

  public async generateScreeningQuestions(
    candidate: { candidateName: string; fitmentScore: number; matchingSkills: string; missingSkills: string; aiSummary: string; hrFeedback: string },
    job: { jobTitle: string; department: string; requiredSkills: string; goodToHaveSkills: string; experience: string }
  ): Promise<string[]> {
    const prompt = `You are a senior technical interviewer. Generate exactly 8 targeted interview questions for this candidate and role.

Role: ${job.jobTitle} — ${job.department}
Experience Required: ${job.experience}
Required Skills: ${job.requiredSkills}
${job.goodToHaveSkills ? `Good To Have: ${job.goodToHaveSkills}` : ''}

Candidate: ${candidate.candidateName}
Fitment Score: ${candidate.fitmentScore}%
Matching Skills: ${candidate.matchingSkills || 'None identified'}
Missing Skills: ${candidate.missingSkills || 'None identified'}
AI Assessment: ${candidate.aiSummary || 'Not available'}
${candidate.hrFeedback ? `HR Notes: ${candidate.hrFeedback}` : ''}

Question criteria:
- 3 questions verifying depth in matching skills (scenario-based, not "do you know X?")
- 3 questions probing missing skills or learning ability to bridge gaps
- 2 behavioral/situational questions specific to this role

Return ONLY a JSON array of 8 strings. No explanation, no numbering, no markdown:
["question one", "question two", ...]`;

    const rawText = await callClaude({
      model: MODEL,
      max_tokens: 1024,
      temperature: 0.5,
      system: 'You are a senior technical interviewer. Return only a valid JSON array of strings. No explanation, no markdown, no text outside the JSON array.',
      messages: [{ role: 'user', content: prompt }],
    });

    const cleaned = rawText.replace(/```json/gi, '').replace(/```/g, '').trim();
    const parsed = JSON.parse(cleaned) as unknown[];
    if (!Array.isArray(parsed)) throw new Error('Expected array from AI');
    return parsed.map(String);
  }

  public async screenResume(
    resume: string,
    jobTitle: string,
    department: string,
    requiredSkills: string,
    experience: string
  ): Promise<IFitmentReport> {
    if (!resume || resume.trim().length === 0) {
      throw new Error('No resume content provided.');
    }

    const userContent = buildUserContent(resume, jobTitle, department, requiredSkills, experience);

    const request: IAnthropicRequest = {
      model: MODEL,
      max_tokens: MAX_TOKENS,
      temperature: 0,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userContent }],
    };

    // First attempt
    let rawText: string;
    try {
      rawText = await callClaude(request);
    } catch (err) {
      throw new Error(`Claude API call failed: ${(err as Error).message}`);
    }

    // Parse — retry once if JSON is malformed
    try {
      return parseReport(rawText);
    } catch {
      console.warn('[AIService] JSON parse failed on first attempt — retrying with fix instruction');
      try {
        const retryRequest: IAnthropicRequest = {
          ...request,
          messages: [
            { role: 'user', content: userContent },
            { role: 'assistant', content: rawText },
            { role: 'user', content: 'Your previous response was not valid JSON. Return only the JSON object, nothing else.' },
          ],
        };
        rawText = await callClaude(retryRequest);
        return parseReport(rawText);
      } catch (retryErr) {
        throw new Error(
          `Failed to parse AI response after retry: ${(retryErr as Error).message}\nRaw: ${rawText}`
        );
      }
    }
  }
}
