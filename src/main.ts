import {randomBytes} from "node:crypto";
import {createServer} from "node:http";
import {
  buildGoogleAuthorizeUrl,
  createGoogleCalendarEvent,
  deleteGoogleCalendarEvent,
  ensureGoogleCalendar,
  exchangeGoogleCodeForToken,
  fetchGoogleCalendarEvents,
  updateGoogleCalendarEvent,
} from "./google.ts";
import {
  buildMicrosoftAuthorizeUrl,
  exchangeMicrosoftCodeForToken,
  fetchMicrosoftTodoListsWithTasks,
} from "./microsoft.ts";
import {createPkce, toBase64Url} from "./utils.ts";
import {z} from "zod";

const configurationSchema = z.object({
  port: z.coerce.number().default(3000),
  ms: z.object({
    tenant: z.string().default("common"),
    clientId: z.string(),
    clientSecret: z.string(),
  }),
  google: z.object({
    clientId: z.string(),
    clientSecret: z.string(),
  }),
});

const configuration = configurationSchema.parse({
  port: process.env.PORT,
  ms: {
    tenant: process.env.MS_TENANT,
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
  },
  google: {
    clientId: process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  },
})

const buildEventTimesFromTaskDue = (task: { dueDateTime?: { dateTime: string; timeZone?: string } }) => {
  if (!task.dueDateTime?.dateTime) {
    return null;
  }

  return {
    start: {
      dateTime: task.dueDateTime.dateTime,
      timeZone: task.dueDateTime.timeZone || "UTC",
    },
    end: {
      dateTime: task.dueDateTime.dateTime,
      timeZone: task.dueDateTime.timeZone || "UTC",
    },
  };
};

const extractTaskIdFromEventDescription = (description?: string) => {
  if (!description) {
    return null;
  }

  const match = description.match(/Task ID:\s*(.+)/);
  return match?.[1]?.trim() || null;
};

const buildEventSummary = (listName: string, taskTitle: string) => `[${listName}] ${taskTitle}`;

const toTimestamp = (value?: string) => {
  if (!value) {
    return null;
  }

  const time = new Date(value).getTime();
  return Number.isNaN(time) ? null : time;
};

const eventNeedsUpdate = (params: {
  currentSummary?: string;
  desiredSummary: string;
  currentStartDateTime?: string;
  currentEndDateTime?: string;
  desiredStartDateTime?: string;
  desiredEndDateTime?: string;
}) => {
  if (params.currentSummary !== params.desiredSummary) {
    return true;
  }

  const currentStart = toTimestamp(params.currentStartDateTime);
  const currentEnd = toTimestamp(params.currentEndDateTime);
  const desiredStart = toTimestamp(params.desiredStartDateTime);
  const desiredEnd = toTimestamp(params.desiredEndDateTime);

  return currentStart !== desiredStart || currentEnd !== desiredEnd;
};

const waitForOAuthCodeWithServer = async (params: {
  port: number;
  host: string;
  baseUrl: string;
  state: string;
  providerName: string;
}) => {
  let handled = false;
  const {promise, resolve, reject} = Promise.withResolvers<string>();

  const server = createServer((req, res) => {
    try {
      const url = new URL(req.url ?? "/", params.baseUrl);

      const code = url.searchParams.get("code");
      const error = url.searchParams.get("error");
      const errorDescription = url.searchParams.get("error_description");
      const returnedState = url.searchParams.get("state");

      if (error) {
        res.writeHead(400, {"Content-Type": "text/plain; charset=utf-8"});
        res.end(`${params.providerName} login error: ${error}\n${errorDescription ?? ""}`);
        server.close();
        reject(new Error(`${params.providerName} OAuth error: ${error} ${errorDescription ?? ""}`));
        return;
      }

      if (!code) {
        res.writeHead(200, {"Content-Type": "text/plain; charset=utf-8"});
        res.end("Callback server działa. Otwórz link logowania z konsoli.");
        return;
      }

      if (handled) {
        res.writeHead(200, {"Content-Type": "text/plain; charset=utf-8"});
        res.end("Authorization code was already processed.");
        return;
      }

      if (returnedState !== params.state) {
        res.writeHead(400, {"Content-Type": "text/plain; charset=utf-8"});
        res.end("Invalid OAuth state.");
        server.close();
        reject(new Error(`${params.providerName} OAuth state mismatch.`));
        return;
      }

      handled = true;
      res.writeHead(200, {"Content-Type": "text/plain; charset=utf-8"});
      res.end(`${params.providerName} login successful. Wróć do konsoli.`);
      server.close();
      resolve(code);
    } catch (error) {
      res.writeHead(500, {"Content-Type": "text/plain; charset=utf-8"});
      res.end("Internal server error. Check console.");
      server.close();
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });

  server.listen(params.port, params.host, () => {
    console.log(`Server started on ${params.baseUrl} (${params.providerName})`);
  });

  return promise.finally(() => server.close());
};

const host = "127.0.0.1";
const baseUrl = `http://localhost:${configuration.port}`;
const msScopes = "https://graph.microsoft.com/Tasks.Read offline_access";

const msPkce = createPkce();
const msState = toBase64Url(randomBytes(16));
const msAuthorizeUrl = buildMicrosoftAuthorizeUrl({
  tenant: configuration.ms.tenant,
  clientId: configuration.ms.clientId,
  redirectUri: baseUrl,
  scope: msScopes,
  challenge: msPkce.challenge,
  state: msState,
});

console.log("1) Zaloguj się do Microsoft (To Do):");
console.log(msAuthorizeUrl.toString());

const msCode = await waitForOAuthCodeWithServer({
  port: configuration.port,
  host,
  baseUrl,
  state: msState,
  providerName: "Microsoft",
});

const msToken = await exchangeMicrosoftCodeForToken({
  authorizationCode: msCode,
  clientId: configuration.ms.clientId,
  clientSecret: configuration.ms.clientSecret,
  tenant: configuration.ms.tenant,
  redirectUri: baseUrl,
  codeVerifier: msPkce.verifier,
  scope: msScopes,
});

console.log("Microsoft token expires in (s):", msToken.expires_in);
console.log("Pobieram listy zadań z Microsoft Graph...");

const todoLists = await fetchMicrosoftTodoListsWithTasks(msToken.access_token);
const googleScopes = "https://www.googleapis.com/auth/calendar";

const googlePkce = createPkce();
const googleState = toBase64Url(randomBytes(16));
const googleAuthorizeUrl = buildGoogleAuthorizeUrl({
  clientId: configuration.google.clientId,
  redirectUri: baseUrl,
  scope: googleScopes,
  challenge: googlePkce.challenge,
  state: googleState,
});

console.log("\n2) Zaloguj się do Google Calendar:");
console.log(googleAuthorizeUrl.toString());

const googleCode = await waitForOAuthCodeWithServer({
  port: configuration.port,
  host,
  baseUrl,
  state: googleState,
  providerName: "Google",
});

const googleToken = await exchangeGoogleCodeForToken({
  authorizationCode: googleCode,
  clientId: configuration.google.clientId,
  clientSecret: configuration.google.clientSecret,
  redirectUri: baseUrl,
  codeVerifier: googlePkce.verifier,
});

console.log("Google token expires in (s):", googleToken.expires_in);
console.log("Pobieram kalendarze i wydarzenia z Google Calendar API...");

const microsoftTodoCalendar = await ensureGoogleCalendar(googleToken.access_token, "Microsoft TODO");

const existingTodoEvents = await fetchGoogleCalendarEvents(
  googleToken.access_token,
  microsoftTodoCalendar.id,
);
const existingEventsByTaskId = new Map(
  existingTodoEvents
    .map((event) => {
      const taskId = extractTaskIdFromEventDescription(event.description);
      return taskId ? [taskId, event] as const : null;
    })
    .filter((entry): entry is readonly [string, (typeof existingTodoEvents)[number]] => Boolean(entry)),
);

let createdEvents = 0;
let updatedEvents = 0;
let deletedEvents = 0;
let unchangedEvents = 0;
let skippedTasksWithoutDue = 0;
let skippedTasksInvalidDue = 0;

for (const {list, tasks} of todoLists) {
  for (const task of tasks) {
    const existingEvent = existingEventsByTaskId.get(task.id);

    if (task.status === "completed") {
      if (existingEvent) {
        await deleteGoogleCalendarEvent(
          googleToken.access_token,
          microsoftTodoCalendar.id,
          existingEvent.id,
        );
        existingEventsByTaskId.delete(task.id);
        deletedEvents += 1;
      }
      continue;
    }

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

    const desiredSummary = buildEventSummary(list.displayName, task.title);
    const desiredDescription = `Microsoft To Do\nLista: ${list.displayName}\nTask ID: ${task.id}`;
    if (existingEvent) {
      if (
        !eventNeedsUpdate({
          currentSummary: existingEvent.summary,
          desiredSummary,
          currentStartDateTime: existingEvent.start?.dateTime,
          currentEndDateTime: existingEvent.end?.dateTime,
          desiredStartDateTime: eventTimes.start.dateTime,
          desiredEndDateTime: eventTimes.end.dateTime,
        })
      ) {
        unchangedEvents += 1;
        continue;
      }

      await updateGoogleCalendarEvent(
        googleToken.access_token,
        microsoftTodoCalendar.id,
        existingEvent.id,
        {
          summary: desiredSummary,
          description: desiredDescription,
          start: eventTimes.start,
          end: eventTimes.end,
        },
      );
      updatedEvents += 1;
      continue;
    }

    await createGoogleCalendarEvent(googleToken.access_token, microsoftTodoCalendar.id, {
      summary: desiredSummary,
      description: desiredDescription,
      start: eventTimes.start,
      end: eventTimes.end,
    });
    existingEventsByTaskId.set(task.id, {
      id: task.id,
      summary: desiredSummary,
      description: desiredDescription,
      start: eventTimes.start,
      end: eventTimes.end,
    });
    createdEvents += 1;
  }
}

console.log(`Utworzono wydarzenia: ${createdEvents}`);
console.log(`Zaktualizowano: ${updatedEvents}`);
console.log(`Usunięto: ${deletedEvents}`);
console.log(`Bez zmian: ${unchangedEvents}`);
console.log(`Pominięto bez terminu: ${skippedTasksWithoutDue}`);
console.log(`Pominięto z błędnym terminem: ${skippedTasksInvalidDue}`);
