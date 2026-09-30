import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IImportantEmail, IMeeting, INewsPost, IPlanMyDayData, IPlannerTask } from './models';

const AGENDA_MAX: number = 10;
const TASKS_MAX: number = 5;
const EMAILS_MAX: number = 5;
const NEWS_MAX: number = 4;

async function getGraphClient(context: ApplicationCustomizerContext): Promise<MSGraphClientV3> {
  return context.msGraphClientFactory.getClient('3');
}

interface IGraphDateTimeZone {
  dateTime: string;
}

interface IGraphEvent {
  id: string;
  subject: string;
  start: IGraphDateTimeZone;
  end: IGraphDateTimeZone;
  isAllDay: boolean;
  onlineMeeting?: { joinUrl?: string };
}

interface IGraphPlannerTask {
  id: string;
  title: string;
  dueDateTime?: string;
  percentComplete: number;
  planId: string;
}

interface IGraphMessage {
  id: string;
  subject: string;
  from?: { emailAddress?: { name?: string } };
  receivedDateTime: string;
  webLink: string;
}

async function getDisplayName(client: MSGraphClientV3): Promise<string> {
  const me: { displayName: string } = await client.api('/me').select('displayName').get();
  return me.displayName;
}

async function getTodayAgenda(client: MSGraphClientV3): Promise<IMeeting[]> {
  const now: Date = new Date();
  const midnight: Date = new Date(now);
  midnight.setHours(23, 59, 59, 999);

  const response: { value: IGraphEvent[] } = await client
    .api('/me/calendarView')
    .header('Prefer', 'outlook.timezone="UTC"')
    .query({
      startDateTime: now.toISOString(),
      endDateTime: midnight.toISOString()
    })
    .select('id,subject,start,end,isAllDay,onlineMeeting')
    .orderby('start/dateTime')
    .top(AGENDA_MAX)
    .get();

  return response.value
    .filter(event => !event.isAllDay)
    .map(event => ({
      id: event.id,
      subject: event.subject,
      start: new Date(`${event.start.dateTime}Z`),
      end: new Date(`${event.end.dateTime}Z`),
      joinUrl: event.onlineMeeting?.joinUrl,
      isAllDay: event.isAllDay
    }));
}

async function getDueTasks(client: MSGraphClientV3): Promise<IPlannerTask[]> {
  const endOfToday: Date = new Date();
  endOfToday.setHours(23, 59, 59, 999);

  const response: { value: IGraphPlannerTask[] } = await client.api('/me/planner/tasks').get();

  const dueOrOverdue: IGraphPlannerTask[] = response.value
    .filter(task => task.percentComplete < 100 && task.dueDateTime && new Date(task.dueDateTime) <= endOfToday)
    .slice(0, TASKS_MAX);

  const planTitleCache: Map<string, string> = new Map<string, string>();
  const tasks: IPlannerTask[] = [];
  for (const task of dueOrOverdue) {
    if (!planTitleCache.has(task.planId)) {
      try {
        const plan: { title: string } = await client.api(`/planner/plans/${task.planId}`).select('title').get();
        planTitleCache.set(task.planId, plan.title);
      } catch {
        planTitleCache.set(task.planId, '');
      }
    }
    tasks.push({
      id: task.id,
      title: task.title,
      dueDateTime: task.dueDateTime ? new Date(task.dueDateTime) : undefined,
      planTitle: planTitleCache.get(task.planId) || '',
      percentComplete: task.percentComplete
    });
  }
  return tasks;
}

async function getImportantEmails(client: MSGraphClientV3): Promise<IImportantEmail[]> {
  const response: { value: IGraphMessage[] } = await client
    .api('/me/messages')
    //.filter("importance eq 'high'")
    .select('id,subject,from,receivedDateTime,webLink')
    .orderby('receivedDateTime desc')
    .top(EMAILS_MAX)
    .get();

  return response.value.map(message => ({
    id: message.id,
    subject: message.subject,
    from: message.from?.emailAddress?.name || '',
    receivedDateTime: new Date(message.receivedDateTime),
    webLink: message.webLink
  }));
}

interface ISearchRow {
  Cells: { Key: string; Value: string }[] | { results: { Key: string; Value: string }[] };
}

interface ISearchResponse {
  PrimaryQueryResult?: {
    RelevantResults?: { Table?: { Rows?: ISearchRow[] | { results: ISearchRow[] } } };
  };
}

// Search REST returns plain arrays with odata=nometadata and { results: [] } wrappers with odata=verbose.
function unwrap<T>(value: T[] | { results: T[] } | undefined): T[] {
  if (!value) {
    return [];
  }
  return Array.isArray(value) ? value : value.results || [];
}

function getCellValue(row: ISearchRow, key: string): string {
  const cell = unwrap(row.Cells).find(c => c.Key === key);
  return cell ? cell.Value : '';
}

async function getLatestNews(context: ApplicationCustomizerContext): Promise<INewsPost[]> {
  const webUrl: string = context.pageContext.web.absoluteUrl;
  const response: SPHttpClientResponse = await context.spHttpClient.post(
    `${webUrl}/_api/search/postquery`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify({
        request: {
          // News posts are pages in the Site Pages library, not a plain document library
          Querytext: 'PromotedState:2 contentclass:STS_ListItem_WebPageLibrary',
          RowLimit: NEWS_MAX,
          SelectProperties: ['Title', 'Path', 'LastModifiedTime', 'PictureThumbnailURL'],
          SortList: [{ Property: 'LastModifiedTime', Direction: 1 }],
          TrimDuplicates: false,
          ClientType: 'ContentSearchRegular'
        }
      })
    }
  );

  if (!response.ok) {
    console.error(`[PlanMyDay] News search failed (${response.status}): ${await response.text()}`);
    return [];
  }

  const json: ISearchResponse = await response.json();
  const rows: ISearchRow[] = unwrap(json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows);

  return rows.map((row, index) => ({
    id: `news-${index}`,
    title: getCellValue(row, 'Title'),
    webUrl: getCellValue(row, 'Path'),
    publishedDateTime: new Date(getCellValue(row, 'LastModifiedTime')),
    imageUrl: getCellValue(row, 'PictureThumbnailURL') || undefined
  }));
}

export async function loadPlanMyDayData(context: ApplicationCustomizerContext): Promise<IPlanMyDayData> {
  const client: MSGraphClientV3 = await getGraphClient(context);

  const [displayName, meetings, tasks, emails, news] = await Promise.all([
    getDisplayName(client),
    getTodayAgenda(client),
    getDueTasks(client),
    getImportantEmails(client),
    getLatestNews(context)
  ]);

  return { displayName, meetings, tasks, emails, news };
}

export async function markTaskComplete(context: ApplicationCustomizerContext, taskId: string): Promise<void> {
  const client: MSGraphClientV3 = await getGraphClient(context);
  const task: { '@odata.etag': string } = await client.api(`/planner/tasks/${taskId}`).select('id').get();
  await client
    .api(`/planner/tasks/${taskId}`)
    .header('If-Match', task['@odata.etag'])
    .patch({ percentComplete: 100 });
}

/**
 * Calls the preview Microsoft 365 Copilot Chat API (`/beta/copilot`). Requires a Copilot
 * add-on license for the signed-in user; the response shape may change since this is a beta API.
 */
export interface CopilotSearchResponse {
  totalCount?: number;
  searchHits?: CopilotSearchHit[];
  "@odata.nextLink"?: string;
}

// Every field is optional: this is a beta API and hits vary by resourceType.
export interface CopilotSearchHit {
  webUrl?: string;
  preview?: string;
  resourceType?: string;
  extracts?: { text?: string }[];
  resourceMetadata?: {
    title?: string;
    author?: string;
  };
}


export async function askCopilot(context: ApplicationCustomizerContext, prompt: string): Promise<CopilotSearchResponse> {
  const client: MSGraphClientV3 = await getGraphClient(context);

  try {
    console.log("[askCopilot] Searching for:", prompt);

    // Build the request body
    const requestBody = {
      query: prompt.trim(),
      pageSize: 10,
      dataSources: {
        oneDrive: {
          resourceMetadataNames: ["title", "author"],
        },
      },
    };

    // POST to /beta/copilot/search
    const response: CopilotSearchResponse =
      await client
        .api("/copilot/search")
        .version("beta")
        .post(requestBody);

    console.log(
      "[askCopilot] Search completed. Total results:",
      response,
    );
    console.log("[askCopilot] Response:", response);

    return response;
  } catch (error) {
    console.error("[askCopilot] Error performing search:", error);
    throw new Error(
      `Failed to perform Copilot search: ${error instanceof Error ? error.message : "Unknown error"}`,
    );
  }

}
