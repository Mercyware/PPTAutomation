# AWS Hosting + CI/CD (MVP)

This guide deploys:

- `server/` to AWS App Runner (`backend` service).
- `addin/` to AWS App Runner (`addin web host` service).
- GitHub Actions CI/CD for both services via ECR images.

## 1. Create ECR repositories

Create two repositories:

- `ppt-automation-server`
- `ppt-automation-addin`

Example:

```bash
aws ecr create-repository --repository-name ppt-automation-server
aws ecr create-repository --repository-name ppt-automation-addin
```

## 2. Create App Runner services

Create two App Runner services from ECR images (initially push once manually or create with placeholder image, then update).

Set service ports:

- backend: `4000`
- addin web host: `3100`

Set runtime environment variables:

### Backend service

- `PORT=4000`
- `NODE_ENV=production`
- LLM provider variables (choose one):
  - Azure OpenAI:
    - `AZURE_OPENAI_ENDPOINT`
    - `AZURE_OPENAI_API_KEY`
    - `AZURE_OPENAI_API_VERSION`
    - `AZURE_OPENAI_CHAT_DEPLOYMENT`
    - `AZURE_OPENAI_CHAT_MODEL`
    - `LLM_TEMPERATURE`
    - `LLM_MAX_TOKENS`
  - Ollama:
    - `OLLAMA_BASE_URL`
    - `OLLAMA_MODEL`

### Addin web host service

- `PORT=3100`
- `NODE_ENV=production`
- `USE_DEV_CERTS=false`
- `BACKEND_BASE_URL=https://<your-backend-service-domain>`

## 3. Configure GitHub OIDC role for deploy

Create an IAM role trusted by GitHub OIDC and scoped to this repository.

Attach permissions for:

- ECR push/pull (`ecr:*` for the two repos, and auth token).
- App Runner deployment trigger (`apprunner:StartDeployment` on both service ARNs).

Store the role ARN in GitHub secret:

- `AWS_ROLE_TO_ASSUME`

## 4. Add GitHub repository variables

Set these repository variables:

- `AWS_REGION`
- `BACKEND_ECR_REPOSITORY` (example: `ppt-automation-server`)
- `ADDIN_ECR_REPOSITORY` (example: `ppt-automation-addin`)
- `BACKEND_APP_RUNNER_SERVICE_ARN`
- `ADDIN_APP_RUNNER_SERVICE_ARN`

## 5. CI/CD workflows

Added workflows:

- `.github/workflows/deploy-backend-aws.yml`
- `.github/workflows/deploy-addin-aws.yml`

Behavior:

- Trigger on push to `main` for matching paths and manual dispatch.
- Build Docker images.
- Push tags `${GITHUB_SHA:0:12}` and `latest` to ECR.
- Trigger App Runner rollout with `aws apprunner start-deployment`.

## 6. Generate hosted manifest for insiders

Generate a hosted manifest with your addin URL:

```powershell
cd addin
.\scripts\generate-hosted-manifest.ps1 `
  -BaseUrl "https://<your-addin-service-domain>" `
  -BackendUrl "https://<your-backend-service-domain>" `
  -OutputPath "manifest.hosted.xml"
```

Share `addin/manifest.hosted.xml` with your insider testers (or deploy via Microsoft 365 admin center).

## 7. Quick verification

Backend:

- `https://<backend-domain>/health`

Addin host:

- `https://<addin-domain>/health`
- `https://<addin-domain>/taskpane.html`

## Notes

- App Runner terminates TLS at the service domain, so the app should run HTTP inside container.
- `addin/server.js` now supports hosted mode via `USE_DEV_CERTS=false`.
- These pipelines are MVP-grade and do not include infra provisioning or rollback automation.
