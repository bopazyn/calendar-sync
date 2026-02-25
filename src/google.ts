import type { TokenResponse } from "./oauth.ts";

type GooglePagedResponse<T> = {
  items?: T[];
  nextPageToken?: string;
};

export type GoogleCalendarListEntry = {
  id: string;
  summary: string;
  primary?: boolean;
  timeZone?: string;
};

export type GoogleCalendarEvent = {
  id: string;
  summary?: string;
  status?: string;
  start?: {
    date?: string;
    dateTime?: string;
    timeZone?: string;
  };
  end?: {
    date?: string;
    dateTime?: string;
    timeZone?: string;
  };
};

export type GoogleCalendarInsertEvent = {
  summary: string;
  description?: string;
  start: {
    date?: string;
    dateTime?: string;
    timeZone?: string;
  };
  end: {
    date?: string;
    dateTime?: string;
    timeZone?: string;
  };
};

export function buildGoogleAuthorizeUrl(params: {
  clientId: string;
  redirectUri: string;
  scope: string;
  challenge: string;
  state: string;
}) {
  const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  url.searchParams.set("client_id", params.clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", params.redirectUri);
  url.searchParams.set("scope", params.scope);
  url.searchParams.set("state", params.state);
  url.searchParams.set("code_challenge", params.challenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("access_type", "offline");
  url.searchParams.set("prompt", "consent");
  return url;
}

export async function exchangeGoogleCodeForToken(params: {
  authorizationCode: string;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  codeVerifier: string;
}) {
  const body = new URLSearchParams({
    client_id: params.clientId,
    client_secret: params.clientSecret,
    code: params.authorizationCode,
    grant_type: "authorization_code",
    redirect_uri: params.redirectUri,
    code_verifier: params.codeVerifier,
  });

  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(`Google token exchange failed: ${res.status} ${JSON.stringify(data)}`);
  }

  return data as TokenResponse;
}

async function googleGet<T>(path: string, accessToken: string) {
  const maxAttempts = 5;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = await fetch(`https://www.googleapis.com${path}`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    const data = await res.json();

    if (res.ok) {
      return data as T;
    }

    if ((res.status === 429 || res.status >= 500) && attempt < maxAttempts) {
      const retryAfterHeader = res.headers.get("retry-after");
      const retryAfterSeconds = retryAfterHeader ? Number(retryAfterHeader) : NaN;
      const waitMs = Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0
        ? retryAfterSeconds * 1000
        : attempt * 1500;

      console.warn(
        `Google API throttling/error (${res.status}) for ${path}. Retry ${attempt}/${maxAttempts - 1} in ${waitMs}ms.`,
      );

      await new Promise((resolve) => setTimeout(resolve, waitMs));
      continue;
    }

    throw new Error(`Google API request failed: ${res.status} ${JSON.stringify(data)}`);
  }

  throw new Error(`Google API request failed after retries for ${path}`);
}

async function googlePost<T>(path: string, accessToken: string, body: unknown) {
  const res = await fetch(`https://www.googleapis.com${path}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(`Google API request failed: ${res.status} ${JSON.stringify(data)}`);
  }

  return data as T;
}

export async function fetchGoogleCalendars(accessToken: string) {
  const items: GoogleCalendarListEntry[] = [];
  let pageToken: string | undefined;

  do {
    const path = new URL("https://www.googleapis.com/calendar/v3/users/me/calendarList");
    if (pageToken) {
      path.searchParams.set("pageToken", pageToken);
    }

    const data = await googleGet<GooglePagedResponse<GoogleCalendarListEntry>>(
      `${path.pathname}${path.search}`,
      accessToken,
    );

    items.push(...(data.items ?? []));
    pageToken = data.nextPageToken;
  } while (pageToken);

  return items;
}

export async function createGoogleCalendar(accessToken: string, summary: string) {
  return await googlePost<GoogleCalendarListEntry>(
    "/calendar/v3/calendars",
    accessToken,
    { summary },
  );
}

export async function ensureGoogleCalendar(accessToken: string, summary: string) {
  const calendars = await fetchGoogleCalendars(accessToken);
  const existing = calendars.find((calendar) => calendar.summary === summary);
  if (existing) {
    return existing;
  }

  return await createGoogleCalendar(accessToken, summary);
}

export async function fetchGoogleCalendarEvents(accessToken: string, calendarId: string) {
  const items: GoogleCalendarEvent[] = [];
  let pageToken: string | undefined;

  do {
    const path = new URL(
      `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events`,
    );
    path.searchParams.set("maxResults", "2500");
    if (pageToken) {
      path.searchParams.set("pageToken", pageToken);
    }

    const data = await googleGet<GooglePagedResponse<GoogleCalendarEvent>>(
      `${path.pathname}${path.search}`,
      accessToken,
    );

    items.push(...(data.items ?? []));
    pageToken = data.nextPageToken;
  } while (pageToken);

  return items;
}

export async function createGoogleCalendarEvent(
  accessToken: string,
  calendarId: string,
  event: GoogleCalendarInsertEvent,
) {
  return await googlePost<GoogleCalendarEvent>(
    `/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events`,
    accessToken,
    event,
  );
}
