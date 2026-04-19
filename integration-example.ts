// Example: How to integrate PPT Microservice with Product Pilot Frontend
// Place this in: frontend/next-app/lib/ppt-service.ts

const PPT_SERVICE_URL = process.env.PPT_SERVICE_URL || 'http://localhost:8000';

interface ProjectStep {
  stepId: number;
  stepName: string;
  data: any;
}

interface GeneratePPTRequest {
  projectName: string;
  projectDescription?: string;
  steps: ProjectStep[];
  template?: 'professional' | 'minimal' | 'dark' | 'startup';
  includeCharts?: boolean;
  includeImages?: boolean;
}

/**
 * Generate PowerPoint presentation via Python microservice
 */
export async function generatePresentation(request: GeneratePPTRequest): Promise<Blob> {
  const response = await fetch(`${PPT_SERVICE_URL}/generate`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(request),
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.detail || 'Failed to generate presentation');
  }

  const result = await response.json();
  
  if (!result.success) {
    throw new Error(result.message);
  }

  // Download the file
  const downloadResponse = await fetch(`${PPT_SERVICE_URL}${result.downloadUrl}`);
  if (!downloadResponse.ok) {
    throw new Error('Failed to download presentation');
  }

  return await downloadResponse.blob();
}

/**
 * List available templates
 */
export async function listTemplates() {
  const response = await fetch(`${PPT_SERVICE_URL}/templates`);
  if (!response.ok) {
    throw new Error('Failed to fetch templates');
  }
  return await response.json();
}

/**
 * Check service health
 */
export async function checkPPTServiceHealth(): Promise<boolean> {
  try {
    const response = await fetch(`${PPT_SERVICE_URL}/health`, {
      method: 'GET',
      cache: 'no-cache',
    });
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Download and trigger browser download of PPT
 */
export function downloadPPT(blob: Blob, filename: string) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.style.display = 'none';
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
  document.body.removeChild(a);
}

// Example usage in a React component:
/*
import { generatePresentation, downloadPPT, checkPPTServiceHealth } from '@/lib/ppt-service';

const handleGeneratePPT = async () => {
  // Check if service is available
  const isHealthy = await checkPPTServiceHealth();
  if (!isHealthy) {
    toast.error('PPT service is unavailable. Using fallback...');
    // Fallback to client-side PptxGenJS
    return;
  }

  try {
    setIsGenerating(true);
    
    const formattedSteps = selectedSteps.map(stepId => ({
      stepId,
      stepName: getStepName(stepId),
      data: getStepData(stepId)
    }));

    const blob = await generatePresentation({
      projectName: selectedProject.name,
      projectDescription: selectedProject.description,
      steps: formattedSteps,
      template: 'professional',
      includeCharts: true,
      includeImages: true
    });

    downloadPPT(blob, `${selectedProject.name}_Presentation.pptx`);
    toast.success('Presentation generated successfully!');
  } catch (error) {
    console.error('PPT generation failed:', error);
    toast.error('Failed to generate presentation');
  } finally {
    setIsGenerating(false);
  }
};
*/
