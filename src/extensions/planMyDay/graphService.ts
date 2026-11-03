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
  onlineMeetingUrl?: string;
  webLink?: string;
}

interface IGraphPlannerTask {
  id: string;
  title: string;
  dueDateTime?: string;
  webUrl?: string;
  percentComplete: number;
  planId: string;
}

interface IGraphCollectionResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

interface IGraphTodoList {
  id: string;
  displayName: string;
}

interface IGraphTodoDateTime {
  dateTime: string;
}

interface IGraphTodoTask {
  id: string;
  title: string;
  status?: string;
  dueDateTime?: IGraphTodoDateTime;
  webLink?: string;
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

function normalizeExternalUrl(value?: string): string | undefined {
  const trimmed: string | undefined = value?.trim();
  if (!trimmed) {
    return undefined;
  }
  return /^https?:\/\//i.test(trimmed) ? trimmed : undefined;
}

function resolveMeetingJoinUrl(event: IGraphEvent): string | undefined {
  return normalizeExternalUrl(event.onlineMeeting?.joinUrl)
    || normalizeExternalUrl(event.onlineMeetingUrl)
    || normalizeExternalUrl(event.webLink);
}

function toGraphApiPath(nextLink: string): string {
  if (/^https?:\/\//i.test(nextLink)) {
    const parsed: URL = new URL(nextLink);
    return `${parsed.pathname}${parsed.search}`;
  }
  return nextLink;
}

function parseGraphDate(dateValue?: string): Date | undefined {
  if (!dateValue) {
    return undefined;
  }
  const parsed: Date = new Date(dateValue);
  return Number.isNaN(parsed.getTime()) ? undefined : parsed;
}

const TODO_ID_PREFIX: string = 'todo:';
const PLANNER_TASKS_WEB_APP_URL: string = 'https://tasks.office.com';
const TODO_WEB_APP_URL: string = 'https://to-do.office.com/tasks/';

function createTodoTaskId(listId: string, taskId: string): string {
  return `${TODO_ID_PREFIX}${encodeURIComponent(listId)}:${encodeURIComponent(taskId)}`;
}

function parseTodoTaskId(taskId: string): { listId: string; taskId: string } | undefined {
  if (!taskId.startsWith(TODO_ID_PREFIX)) {
    return undefined;
  }
  const composite: string = taskId.substring(TODO_ID_PREFIX.length);
  const separatorIndex: number = composite.indexOf(':');
  if (separatorIndex < 0) {
    return undefined;
  }
  const encodedListId: string = composite.substring(0, separatorIndex);
  const encodedTaskId: string = composite.substring(separatorIndex + 1);
  if (!encodedListId || !encodedTaskId) {
    return undefined;
  }
  return { listId: decodeURIComponent(encodedListId), taskId: decodeURIComponent(encodedTaskId) };
}

function sortTasks(tasks: IPlannerTask[]): IPlannerTask[] {
  return tasks.sort((left, right) => {
    const leftDueTime: number = left.dueDateTime?.getTime() ?? Number.MAX_SAFE_INTEGER;
    const rightDueTime: number = right.dueDateTime?.getTime() ?? Number.MAX_SAFE_INTEGER;
    if (leftDueTime !== rightDueTime) {
      return leftDueTime - rightDueTime;
    }
    return left.title.localeCompare(right.title);
  });
}

function getPlannerTaskWebUrl(taskId: string, webUrl?: string): string {
  return normalizeExternalUrl(webUrl) || `${PLANNER_TASKS_WEB_APP_URL}/?taskId=${encodeURIComponent(taskId)}`;
}

function getTodoTaskWebUrl(taskId: string, webLink?: string): string {
  return normalizeExternalUrl(webLink) || `${TODO_WEB_APP_URL}id/${encodeURIComponent(taskId)}`;
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
    .select('id,subject,start,end,isAllDay,onlineMeeting,onlineMeetingUrl,webLink')
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
      joinUrl: resolveMeetingJoinUrl(event),
      isAllDay: event.isAllDay
    }));
}

async function getDueTasks(client: MSGraphClientV3): Promise<IPlannerTask[]> {
  const endOfToday: Date = new Date();
  endOfToday.setHours(23, 59, 59, 999);

  const plannerTasks: IGraphPlannerTask[] = [];
  let nextLink: string | undefined = '/me/planner/tasks?$select=id,title,dueDateTime,webUrl,percentComplete,planId&$top=100';

  while (nextLink) {
    const page: IGraphCollectionResponse<IGraphPlannerTask> = await client.api(toGraphApiPath(nextLink)).get();
    plannerTasks.push(...page.value);
    nextLink = page['@odata.nextLink'];
  }

  const planTitleCache: Map<string, string> = new Map<string, string>();
  const plannerOpenTasks: IPlannerTask[] = [];
  for (const task of plannerTasks.filter(value => (value.percentComplete ?? 0) < 100)) {
    if (!planTitleCache.has(task.planId)) {
      try {
        const plan: { title: string } = await client.api(`/planner/plans/${task.planId}`).select('title').get();
        planTitleCache.set(task.planId, plan.title);
      } catch {
        planTitleCache.set(task.planId, '');
      }
    }
    plannerOpenTasks.push({
      id: task.id,
      title: task.title,
      dueDateTime: parseGraphDate(task.dueDateTime),
      webUrl: getPlannerTaskWebUrl(task.id, task.webUrl),
      planTitle: planTitleCache.get(task.planId) || '',
      percentComplete: task.percentComplete ?? 0
    });
  }

  const todoOpenTasks: IPlannerTask[] = [];
  try {
    const todoLists: IGraphTodoList[] = [];
    let todoListsNextLink: string | undefined = '/me/todo/lists';
    while (todoListsNextLink) {
      const todoListsPage: IGraphCollectionResponse<IGraphTodoList> = await client.api(toGraphApiPath(todoListsNextLink)).get();
      todoLists.push(...todoListsPage.value);
      todoListsNextLink = todoListsPage['@odata.nextLink'];
    }
    const todoListResults: Array<{ listId: string; listTitle: string; task: IGraphTodoTask }> = [];

    await Promise.all(todoLists.map(async (list: IGraphTodoList) => {
      let listNextLink: string | undefined = `/me/todo/lists/${encodeURIComponent(list.id)}/tasks`;
      while (listNextLink) {
        const listPage: IGraphCollectionResponse<IGraphTodoTask> = await client.api(toGraphApiPath(listNextLink)).get();
        listPage.value
          .filter(task => (task.status || '').toLowerCase() !== 'completed')
          .forEach(task => todoListResults.push({ listId: list.id, listTitle: list.displayName, task }));
        listNextLink = listPage['@odata.nextLink'];
      }
    }));

    todoListResults.forEach(({ listId, listTitle, task }) => {
      todoOpenTasks.push({
        id: createTodoTaskId(listId, task.id),
        title: task.title,
        dueDateTime: parseGraphDate(task.dueDateTime?.dateTime),
        webUrl: getTodoTaskWebUrl(task.id, task.webLink),
        planTitle: listTitle,
        percentComplete: 0
      });
    });
  } catch (error) {
    console.warn('[PlanMyDay] To Do tasks unavailable, continuing with Planner tasks only.', error);
  }

  const openTasks: IPlannerTask[] = [...plannerOpenTasks, ...todoOpenTasks];
  const dueOrOverdue: IPlannerTask[] = sortTasks(
    openTasks.filter(task => Boolean(task.dueDateTime && task.dueDateTime.getTime() <= endOfToday.getTime()))
  );

  const selectedTasks: IPlannerTask[] = sortTasks(dueOrOverdue.length > 0 ? dueOrOverdue : openTasks).slice(0, TASKS_MAX);
  return selectedTasks;
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
  const todoTask: { listId: string; taskId: string } | undefined = parseTodoTaskId(taskId);
  if (todoTask) {
    await client
      .api(`/me/todo/lists/${encodeURIComponent(todoTask.listId)}/tasks/${encodeURIComponent(todoTask.taskId)}`)
      .patch({ status: 'completed' });
    return;
  }
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

interface IGraphSearchResource {
  webUrl?: string;
  name?: string;
  title?: string;
}

interface IGraphSearchHit {
  summary?: string;
  resource?: IGraphSearchResource;
}

interface IGraphSearchHitsContainer {
  total?: number;
  hits?: IGraphSearchHit[];
}

interface IGraphSearchResponse {
  value?: Array<{
    hitsContainers?: IGraphSearchHitsContainer[];
  }>;
}

function mapGraphSearchToCopilot(response: IGraphSearchResponse): CopilotSearchResponse {
  const container: IGraphSearchHitsContainer | undefined = response.value?.[0]?.hitsContainers?.[0];
  const hits: IGraphSearchHit[] = container?.hits || [];
  return {
    totalCount: container?.total ?? hits.length,
    searchHits: hits.map(hit => ({
      webUrl: hit.resource?.webUrl,
      preview: hit.summary,
      resourceType: 'searchResult',
      resourceMetadata: {
        title: hit.resource?.name || hit.resource?.title
      }
    }))
  };
}


export async function askCopilot(context: ApplicationCustomizerContext, prompt: string): Promise<CopilotSearchResponse> {
  const client: MSGraphClientV3 = await getGraphClient(context);

  try {
    console.log("[askCopilot] Searching for:", prompt);

    const normalizedPrompt: string = prompt.trim();

    // Build broad request body (no datasource restriction)
    const requestBody = {
      query: normalizedPrompt,
      pageSize: 10
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
    const hasHits: boolean = Array.isArray(response.searchHits) && response.searchHits.length > 0;
    if (hasHits) {
      console.log("[askCopilot] Response:", response);
      return response;
    }

    console.log('[askCopilot] Copilot search returned no hits, falling back to Graph Search API.');
    const graphSearchResponse: IGraphSearchResponse = await client
      .api('/search/query')
      .version('v1.0')
      .post({
        requests: [
          {
            entityTypes: ['driveItem', 'listItem', 'site'],
            query: { queryString: normalizedPrompt },
            from: 0,
            size: 10
          }
        ]
      });

    const fallbackResponse: CopilotSearchResponse = mapGraphSearchToCopilot(graphSearchResponse);
    console.log('[askCopilot] Fallback search response:', fallbackResponse);
    return fallbackResponse;
  } catch (error) {
    console.error("[askCopilot] Error performing search:", error);
    throw new Error(
      `Failed to perform Copilot search: ${error instanceof Error ? error.message : "Unknown error"}`,
    );
  }

}
