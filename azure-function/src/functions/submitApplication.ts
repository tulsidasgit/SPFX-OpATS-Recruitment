import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { corsHeaders } from '../shared/spClient';
import { submitApplication, IApplicationPayload } from './formWriter';

const ALLOWED_EXTENSIONS = ['.pdf', '.docx', '.doc'];
const MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024; // 5 MB

function validatePayload(body: Partial<IApplicationPayload>): string | null {
  if (!body.jobId || isNaN(Number(body.jobId)))  return 'jobId is required and must be a number.';
  if (!body.name?.trim())                         return 'Full name is required.';
  if (!body.email?.trim())                        return 'Email address is required.';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(body.email)) return 'Please enter a valid email address.';
  if (!body.phone?.trim())                        return 'Contact number is required.';
  if (!body.currentCtc?.trim())                   return 'Current CTC is required.';
  if (!body.expectedCtc?.trim())                  return 'Expected CTC is required.';
  if (!body.noticePeriod?.trim())                 return 'Notice period is required.';
  if (!body.resumeBase64?.trim())                 return 'Resume file is required.';
  if (!body.resumeFileName?.trim())               return 'Resume file name is required.';

  const ext = ('.' + body.resumeFileName.split('.').pop()).toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    return `Only PDF and DOCX resumes are accepted. Received: ${ext}`;
  }

  const sizeBytes = Math.ceil((body.resumeBase64.length * 3) / 4);
  if (sizeBytes > MAX_FILE_SIZE_BYTES) {
    return 'Resume file exceeds the 5 MB limit. Please compress and re-upload.';
  }

  return null;
}

async function submitApplicationHandler(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const headers = corsHeaders();

  if (request.method === 'OPTIONS') {
    return { status: 204, headers };
  }

  let body: Partial<IApplicationPayload>;
  try {
    body = (await request.json()) as Partial<IApplicationPayload>;
  } catch {
    return { status: 400, headers, jsonBody: { error: 'Invalid JSON body.' } };
  }

  const validationError = validatePayload(body);
  if (validationError) {
    return { status: 400, headers, jsonBody: { error: validationError } };
  }

  try {
    const result = await submitApplication(body as IApplicationPayload);

    context.log(
      `SubmitApplication: candidate #${result.candidateId} created for job #${body.jobId} — ${body.name}`
    );

    return {
      status: 200,
      headers,
      jsonBody: {
        success: true,
        candidateId: result.candidateId,
        message: 'Application submitted successfully. We will be in touch soon.',
      },
    };
  } catch (err) {
    const message = (err as Error).message ?? 'Unknown error';
    context.error('SubmitApplication error:', message);

    // Surface known user-facing errors directly; hide internal ones
    const isUserFacing =
      message.includes('not found') ||
      message.includes('no longer accepting') ||
      message.includes('Only PDF');

    return {
      status: isUserFacing ? 400 : 500,
      headers,
      jsonBody: {
        error: isUserFacing
          ? message
          : 'Failed to submit your application. Please try again or contact us directly.',
      },
    };
  }
}

app.http('SubmitApplication', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: submitApplicationHandler,
});
