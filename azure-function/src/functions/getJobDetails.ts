import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getSpClient, corsHeaders } from '../shared/spClient';

interface IJobDetailsResponse {
  id: number;
  title: string;
  department: string;
  jobTitle: string;
  jobLocation: string;
  jobType: string;
  experience: string;
  requiredSkills: string;
  goodToHaveSkills: string;
  jobDescription: string;
}

async function getJobDetails(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  const headers = corsHeaders();

  // Handle CORS preflight
  if (request.method === 'OPTIONS') {
    return { status: 204, headers };
  }

  const jobIdParam = request.query.get('jobId');
  if (!jobIdParam || isNaN(Number(jobIdParam))) {
    return {
      status: 400,
      headers,
      jsonBody: { error: 'Missing or invalid jobId query parameter.' },
    };
  }

  const jobId = Number(jobIdParam);

  try {
    const sp = getSpClient();

    const items = await sp.web.lists
      .getByTitle('JobOpenings')
      .items.select(
        'Id', 'Title', 'Department', 'JobTitle', 'JobLocation', 'JobType',
        'Experience', 'RequiredSkills', 'GoodToHaveSkills', 'JobDescription', 'Status'
      )
      .filter(`Id eq ${jobId}`)
      .top(1)();

    if (items.length === 0) {
      return {
        status: 404,
        headers,
        jsonBody: { error: `Job opening #${jobId} not found.` },
      };
    }

    const item = items[0] as Record<string, unknown>;

    if (item['Status'] === 'Closed') {
      return {
        status: 410,
        headers,
        jsonBody: { error: 'This job opening is no longer accepting applications.' },
      };
    }

    const response: IJobDetailsResponse = {
      id: item['Id'] as number,
      title: (item['Title'] as string) ?? '',
      department: (item['Department'] as string) ?? '',
      jobTitle: (item['JobTitle'] as string) ?? '',
      jobLocation: (item['JobLocation'] as string) ?? '',
      jobType: (item['JobType'] as string) ?? '',
      experience: (item['Experience'] as string) ?? '',
      requiredSkills: (item['RequiredSkills'] as string) ?? '',
      goodToHaveSkills: (item['GoodToHaveSkills'] as string) ?? '',
      jobDescription: (item['JobDescription'] as string) ?? '',
    };

    context.log(`GetJobDetails: served job #${jobId} — "${response.title}"`);

    return { status: 200, headers, jsonBody: response };
  } catch (err) {
    context.error('GetJobDetails error:', err);
    return {
      status: 500,
      headers,
      jsonBody: { error: 'Failed to retrieve job details. Please try again later.' },
    };
  }
}

app.http('GetJobDetails', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: getJobDetails,
});
