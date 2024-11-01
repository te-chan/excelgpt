import { BrainCircuit, Router } from "lucide-react";

export interface ProviderInput {
  id: string;
  label: string;
  placeholder: string;
  icon: React.ReactNode;
  validation?: (value: string) => string | null;
}

export interface AIProviderConfig {
  name: string;
  inputs: ProviderInput[];
}

export const aiProviders: AIProviderConfig[] = [
  {
    name: 'OpenAI',
    inputs: [],
  },
  {
    name: 'Azure OpenAI',
    inputs: [
      {
        id: 'azureDeployment',
        label: 'デプロイ名',
        placeholder: 'gpt-8-human',
        icon: <BrainCircuit className="w-4 h-4" />,
      },
      {
        id: 'azureEndpoint',
        label: 'APIエンドポイント',
        placeholder: 'https://deployment.azure.somewhere',
        icon: <Router className="w-4 h-4" />,
      },
    ],
  },
  // 他のプロバイダーをここに追加
];