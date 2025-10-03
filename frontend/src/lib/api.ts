import axios from 'axios';

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000';

export interface UploadResponse {
  file_id: string;
  filename: string;
  path: string;
}

export interface TranslationJob {
  id: string;
  status: 'pending' | 'running' | 'completed' | 'failed';
  filename: string;
  model: string;
  progress: number;
  created_at: string;
  output_file?: string;
  error?: string;
}

export const api = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json',
  },
});

export const uploadFile = async (file: File): Promise<UploadResponse> => {
  const formData = new FormData();
  formData.append('file', file);

  const response = await api.post<UploadResponse>('/upload', formData, {
    headers: {
      'Content-Type': 'multipart/form-data',
    },
  });

  return response.data;
};

export const startTranslation = async (
  fileId: string,
  filename: string,
  model: string = 'gpt-4o'
): Promise<{ job_id: string; status: string }> => {
  const response = await api.post<{ job_id: string; status: string }>('/translate', {
    file_id: fileId,
    filename,
    model,
  });

  return response.data;
};

export const getJobStatus = async (jobId: string): Promise<TranslationJob> => {
  const response = await api.get<TranslationJob>(`/jobs/${jobId}`);
  return response.data;
};

export const downloadResult = async (jobId: string): Promise<Blob> => {
  const response = await api.get(`/jobs/${jobId}/download`, {
    responseType: 'blob',
  });
  return response.data;
};

export const listJobs = async (): Promise<TranslationJob[]> => {
  const response = await api.get<TranslationJob[]>('/jobs');
  return response.data;
};