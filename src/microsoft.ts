import type { TokenResponse } from "./oauth.ts";

type GraphListResponse<T> = {
  value: T[];
};

export type TodoTaskList = {
  id: string;
  displayName: string;
  wellknownListName?: string;
};

export type TodoTask = {
  id: string;
  title: string;
  status?: string;
  dueDateTime?: {
    dateTime: string;
    timeZone?: string;
  };
};

export type TodoListWithTasks = {
  list: TodoTaskList;
  tasks: TodoTask[];
};

export const buildMicrosoftAuthorizeUrl = (params: {
  tenant: string;
  clientId: string;
  redirectUri: string;
  scope: string;
  challenge: string;
  state: string;
}) => {
  const url = new URL(`https://login.microsoftonline.com/${params.tenant}/oauth2/v2.0/authorize`);
  url.searchParams.set("client_id", params.clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", params.redirectUri);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", params.scope);
  url.searchParams.set("code_challenge", params.challenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("state", params.state);
  return url;
};

export const exchangeMicrosoftCodeForToken = async (params: {
  authorizationCode: string;
  clientId: string;
  tenant: string;
  redirectUri: string;
  codeVerifier: string;
  scope: string;
  clientSecret: string;
}) => {
  const tokenUrl = `https://login.microsoftonline.com/${params.tenant}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: params.clientId,
    grant_type: "authorization_code",
    code: params.authorizationCode,
    redirect_uri: params.redirectUri,
    code_verifier: params.codeVerifier,
    scope: params.scope,
    client_secret: params.clientSecret,
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(`Microsoft token exchange failed: ${res.status} ${JSON.stringify(data)}`);
  }

  return data as TokenResponse;
};

const graphGet = async <T>(path: string, accessToken: string) => {
  const maxAttempts = 5;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    const data = await res.json();

    if (res.ok) {
      return data as T;
    }

    if (res.status === 429 && attempt < maxAttempts) {
      const retryAfterHeader = res.headers.get("retry-after");
      const retryAfterSeconds = retryAfterHeader ? Number(retryAfterHeader) : NaN;
      const waitMs = Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0
        ? retryAfterSeconds * 1000
        : attempt * 1500;

      console.warn(
        `Graph throttling (429) for ${path}. Retry ${attempt}/${maxAttempts - 1} in ${waitMs}ms.`,
      );

      await new Promise((resolve) => setTimeout(resolve, waitMs));
      continue;
    }

    throw new Error(`Graph request failed: ${res.status} ${JSON.stringify(data)}`);
  }

  throw new Error(`Graph request failed after retries for ${path}`);
};

export const fetchMicrosoftTodoListsWithTasks = async (accessToken: string) => {
  const lists = await graphGet<GraphListResponse<TodoTaskList>>("/me/todo/lists", accessToken);

  const tasksByList: TodoListWithTasks[] = [];
  for (const list of lists.value) {
    const tasks = await graphGet<GraphListResponse<TodoTask>>(
      `/me/todo/lists/${list.id}/tasks`,
      accessToken,
    );
    tasksByList.push({ list, tasks: tasks.value });
  }

  return tasksByList;
};
