import { IFitmentReport } from './models';
import { ANTHROPIC_API_KEY } from './AIConfig';

const MODEL = 'claude-haiku-4-5-20251001';
const MAX_TOKENS = 1024;
const MIN_RESUME_LENGTH = 100;
const API_URL = 'https://api.anthropic.com/v1/messages';

const SYSTEM_PROMPT = `You are an expert recruitment AI assistant. Analyse the provided resume \
against the job requirements. Return ONLY valid JSON with absolutely no explanation, \
no markdown formatting, and no code blocks. Be precise and objective about skill matching.`;

interface IAnthropicMessage {
  role: 'user' | 'assistant';
  content: string;
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

function buildUserPrompt(
  jobTitle: string,
  department: string,
  requiredSkills: string,
  experience: string,
  resumeText: string
): string {
  return `Job Title: ${jobTitle}
Department: ${department}
Required Skills: ${requiredSkills}
Experience Required: ${experience}

Resume Text:
${resumeText}

Return exactly this JSON structure with no other text:
{
  "candidateName": "full name extracted from resume",
  "contactEmail": "email address from resume",
  "contactPhone": "phone number from resume",
  "matchingSkills": ["skill1", "skill2"],
  "missingSkills": ["skill1", "skill2"],
  "fitmentScore": 75,
  "experienceMatch": "meets",
  "summary": "2-3 sentence objective assessment of this candidate",
  "recommendation": "Recommended"
}`;
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

  public async screenResume(
    resumeText: string,
    jobTitle: string,
    department: string,
    requiredSkills: string,
    experience: string
  ): Promise<IFitmentReport> {
    if (!resumeText || resumeText.trim().length < MIN_RESUME_LENGTH) {
      throw new Error(
        `Resume text is too short (${resumeText?.trim().length ?? 0} characters). ` +
        `Minimum required: ${MIN_RESUME_LENGTH} characters.`
      );
    }

    const userPrompt = buildUserPrompt(jobTitle, department, requiredSkills, experience, resumeText);

    const request: IAnthropicRequest = {
      model: MODEL,
      max_tokens: MAX_TOKENS,
      temperature: 0,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userPrompt }],
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
            { role: 'user', content: userPrompt },
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
