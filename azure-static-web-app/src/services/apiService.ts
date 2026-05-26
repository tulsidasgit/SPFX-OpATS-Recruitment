export interface IJobDetails {
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

export interface IApplicationPayload {
  jobId: number;
  name: string;
  email: string;
  phone: string;
  currentCtc: string;
  expectedCtc: string;
  noticePeriod: string;
  resumeBase64: string;
  resumeFileName: string;
  gender?: string;
  reasonForLeaving?: string;
  workCultureExpectation?: string;
  linkedInUrl?: string;
}

export interface ISubmitResult {
  success: boolean;
  candidateId: number;
  message: string;
}

const API_BASE = import.meta.env.VITE_API_BASE_URL ?? '';

export async function fetchJobDetails(jobId: string): Promise<IJobDetails> {
  const res = await fetch(`${API_BASE}/api/GetJobDetails?jobId=${encodeURIComponent(jobId)}`);
  const data = await res.json() as IJobDetails & { error?: string };
  if (!res.ok) throw new Error(data.error ?? `HTTP ${res.status}`);
  return data;
}

export async function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.split(',')[1] ?? '');
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
  });
}

export async function submitApplication(payload: IApplicationPayload): Promise<ISubmitResult> {
  const res = await fetch(`${API_BASE}/api/SubmitApplication`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json() as ISubmitResult & { error?: string };
  if (!res.ok) throw new Error(data.error ?? `HTTP ${res.status}`);
  return data;
}
