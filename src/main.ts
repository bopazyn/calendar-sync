import { randomBytes } from "node:crypto";
import { createServer } from "node:http";
import {
  buildGoogleAuthorizeUrl,
  createGoogleCalendarEvent,
  ensureGoogleCalendar,
  exchangeGoogleCodeForToken,
  fetchGoogleCalendarEvents,
  fetchGoogleCalendars,
} from "./google.ts";
import {
  buildMicrosoftAuthorizeUrl,
  exchangeMicrosoftCodeForToken,
  fetchMicrosoftTodoListsWithTasks,
} from "./microsoft.ts";
import { createPkce, toBase64Url } from "./oauth.ts";

function buildEventTimesFromTaskDue(task: { dueDateTime?: { dateTime: string; timeZone?: string } }) {
  if (!task.dueDateTime?.dateTime) {
    return null;
  }

  const startDate = new Date(task.dueDateTime.dateTime);
  if (Number.isNaN(startDate.getTime())) {
    return null;
  }

  const endDate = new Date(startDate.getTime() + 30 * 60 * 1000);

  return {
    start: {
      dateTime: startDate.toISOString(),
      timeZone: task.dueDateTime.timeZone || "UTC",
    },
    end: {
      dateTime: endDate.toISOString(),
      timeZone: task.dueDateTime.timeZone || "UTC",
    },
  };
}

function extractTaskIdFromEventDescription(description?: string) {
  if (!description) {
    return null;
  }

  const match = description.match(/Task ID:\s*(.+)/);
  return match?.[1]?.trim() || null;
}

async function waitForOAuthCodeWithServer(params: {
  port: number;
  host: string;
  baseUrl: string;
  state: string;
  providerName: string;
}) {
  let handled = false;

  return await new Promise<string>((resolve, reject) => {
    const server = createServer((req, res) => {
      try {
        const url = new URL(req.url ?? "/", params.baseUrl);

        const code = url.searchParams.get("code");
        const error = url.searchParams.get("error");
        const errorDescription = url.searchParams.get("error_description");
        const returnedState = url.searchParams.get("state");

        if (error) {
          res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
          res.end(`${params.providerName} login error: ${error}\n${errorDescription ?? ""}`);
          server.close();
          reject(new Error(`${params.providerName} OAuth error: ${error} ${errorDescription ?? ""}`));
          return;
        }

        if (!code) {
          res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
          res.end("Callback server działa. Otwórz link logowania z konsoli.");
          return;
        }

        if (handled) {
          res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
          res.end("Authorization code was already processed.");
          return;
        }

        if (returnedState !== params.state) {
          res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
          res.end("Invalid OAuth state.");
          server.close();
          reject(new Error(`${params.providerName} OAuth state mismatch.`));
          return;
        }

        handled = true;
        res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
        res.end(`${params.providerName} login successful. Wróć do konsoli.`);
        server.close();
        resolve(code);
      } catch (error) {
        res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
        res.end("Internal server error. Check console.");
        server.close();
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });

    server.listen(params.port, params.host, () => {
      console.log(`Server started on ${params.baseUrl} (${params.providerName})`);
    });
  });
}

async function main() {
  const port = Number(process.env.PORT ?? "3000");
  const host = "127.0.0.1";
  const baseUrl = `http://localhost:${port}`;
  const msTenant = process.env.MS_TENANT ?? "common";
  const msClientId = process.env.MS_CLIENT_ID ?? "";
  const msClientSecret = process.env.MS_CLIENT_SECRET ?? "";
  const msScopes = (
    process.env.MS_SCOPES ??
    "https://graph.microsoft.com/Tasks.Read offline_access"
  ).trim();
  const msRedirectUri = process.env.MS_REDIRECT_URI ?? `${baseUrl}/callback/microsoft`;

  if (!msClientId) {
    throw new Error("Missing MS_CLIENT_ID in environment (.env).");
  }
  if (!msClientSecret) {
    throw new Error("Missing MS_CLIENT_SECRET in environment (.env).");
  }

  const msPkce = createPkce();
  const msState = toBase64Url(randomBytes(16));
  const msAuthorizeUrl = buildMicrosoftAuthorizeUrl({
    tenant: msTenant,
    clientId: msClientId,
    redirectUri: msRedirectUri,
    scope: msScopes,
    challenge: msPkce.challenge,
    state: msState,
  });

  console.log("1) Zaloguj się do Microsoft (To Do):");
  console.log(msAuthorizeUrl.toString());

  const msCode = await waitForOAuthCodeWithServer({
    port,
    host,
    baseUrl,
    state: msState,
    providerName: "Microsoft",
  });

  const msToken = await exchangeMicrosoftCodeForToken({
    authorizationCode: msCode,
    clientId: msClientId,
    tenant: msTenant,
    redirectUri: msRedirectUri,
    codeVerifier: msPkce.verifier,
    scope: msScopes,
    clientSecret: msClientSecret,
  });

  console.log("Microsoft token expires in (s):", msToken.expires_in);
  console.log("Pobieram listy zadań z Microsoft Graph...");

  const todoLists = await fetchMicrosoftTodoListsWithTasks(msToken.access_token);
  for (const { list, tasks } of todoLists) {
    console.log(`\nLista: ${list.displayName} (${list.id})`);
    if (tasks.length === 0) {
      console.log("  - brak zadań");
      continue;
    }

    for (const task of tasks) {
      console.log(`  - ${task.title} [${task.status ?? "unknown"}]`);
    }
  }

  const googleClientId = process.env.GOOGLE_CLIENT_ID ?? "";
  const googleClientSecret = process.env.GOOGLE_CLIENT_SECRET ?? "";
  const googleScopes = (
    process.env.GOOGLE_SCOPES ??
    "https://www.googleapis.com/auth/calendar"
  ).trim();
  const googleRedirectUri = process.env.GOOGLE_REDIRECT_URI ?? baseUrl;

  if (!googleClientId) {
    throw new Error("Missing GOOGLE_CLIENT_ID in environment (.env).");
  }
  if (!googleClientSecret) {
    throw new Error("Missing GOOGLE_CLIENT_SECRET in environment (.env).");
  }

  const googlePkce = createPkce();
  const googleState = toBase64Url(randomBytes(16));
  const googleAuthorizeUrl = buildGoogleAuthorizeUrl({
    clientId: googleClientId,
    redirectUri: googleRedirectUri,
    scope: googleScopes,
    challenge: googlePkce.challenge,
    state: googleState,
  });

  console.log("\n2) Zaloguj się do Google Calendar:");
  console.log(googleAuthorizeUrl.toString());

  const googleCode = await waitForOAuthCodeWithServer({
    port,
    host,
    baseUrl,
    state: googleState,
    providerName: "Google",
  });

  const googleToken = await exchangeGoogleCodeForToken({
    authorizationCode: googleCode,
    clientId: googleClientId,
    clientSecret: googleClientSecret,
    redirectUri: googleRedirectUri,
    codeVerifier: googlePkce.verifier,
  });

  console.log("Google token expires in (s):", googleToken.expires_in);
  console.log("Pobieram kalendarze i wydarzenia z Google Calendar API...");

  const calendars = await fetchGoogleCalendars(googleToken.access_token);
  for (const calendar of calendars) {
    console.log(
      `\nKalendarz: ${calendar.summary} (${calendar.id})${calendar.primary ? " [primary]" : ""}`,
    );

    const events = await fetchGoogleCalendarEvents(googleToken.access_token, calendar.id);
    if (events.length === 0) {
      console.log("  - brak wydarzeń");
      continue;
    }

    for (const event of events) {
      const start = event.start?.dateTime ?? event.start?.date ?? "brak daty";
      console.log(`  - ${event.summary ?? "(bez tytułu)"} [${event.status ?? "unknown"}] @ ${start}`);
    }
  }

  const microsoftTodoCalendar = await ensureGoogleCalendar(googleToken.access_token, "Microsoft TODO");
  console.log(`\nKalendarz docelowy: ${microsoftTodoCalendar.summary} (${microsoftTodoCalendar.id})`);

  const existingTodoEvents = await fetchGoogleCalendarEvents(
    googleToken.access_token,
    microsoftTodoCalendar.id,
  );
  const existingTaskIds = new Set(
    existingTodoEvents
      .map((event) => extractTaskIdFromEventDescription(event.description))
      .filter((taskId): taskId is string => Boolean(taskId)),
  );

  let createdEvents = 0;
  let skippedTasksWithoutDue = 0;
  let skippedTasksInvalidDue = 0;
  let skippedTasksDuplicate = 0;

  for (const { list, tasks } of todoLists) {
    for (const task of tasks) {
      if (!task.dueDateTime) {
        skippedTasksWithoutDue += 1;
        continue;
      }

      const eventTimes = buildEventTimesFromTaskDue(task);
      if (!eventTimes) {
        skippedTasksInvalidDue += 1;
        console.warn(`Pomijam zadanie z nieprawidłowym terminem: ${task.title} (${task.id})`);
        continue;
      }

      if (existingTaskIds.has(task.id)) {
        skippedTasksDuplicate += 1;
        continue;
      }

      await createGoogleCalendarEvent(googleToken.access_token, microsoftTodoCalendar.id, {
        summary: task.title,
        description: `Microsoft To Do\nLista: ${list.displayName}\nTask ID: ${task.id}`,
        start: eventTimes.start,
        end: eventTimes.end,
      });
      existingTaskIds.add(task.id);
      createdEvents += 1;
    }
  }

  console.log(
    `\nUtworzono wydarzenia: ${createdEvents}. Pominięto duplikaty: ${skippedTasksDuplicate}. Pominięto bez terminu: ${skippedTasksWithoutDue}. Pominięto z błędnym terminem: ${skippedTasksInvalidDue}.`,
  );
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
