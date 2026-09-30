import { IPlanMyDayData } from './models';

/** Fixed sample data used in the local workbench (?mockData=true), where there is no real Graph context. */
export function getMockData(): IPlanMyDayData {
  const now: Date = new Date();
  const inMinutes = (minutes: number): Date => new Date(now.getTime() + minutes * 60000);

  return {
    displayName: 'Alex Wilber',
    meetings: [
      {
        id: 'mock-meeting-1',
        subject: 'Weekly design sync',
        start: inMinutes(45),
        end: inMinutes(75),
        joinUrl: 'https://teams.microsoft.com/l/meetup-join/mock',
        isAllDay: false
      },
      {
        id: 'mock-meeting-2',
        subject: 'Customer roadmap review',
        start: inMinutes(180),
        end: inMinutes(240),
        joinUrl: 'https://teams.microsoft.com/l/meetup-join/mock2',
        isAllDay: false
      }
    ],
    tasks: [
      { id: 'mock-task-1', title: 'Finish Q3 budget review', dueDateTime: now, planTitle: 'Finance', percentComplete: 0 },
      { id: 'mock-task-2', title: 'Review PR for extension', dueDateTime: inMinutes(-120), planTitle: 'Engineering', percentComplete: 50 }
    ],
    emails: [
      { id: 'mock-email-1', subject: 'Action needed: contract renewal', from: 'Jane Cooper', receivedDateTime: inMinutes(-60), webLink: 'https://outlook.office.com/mail' },
      { id: 'mock-email-2', subject: 'Urgent: production incident summary', from: 'IT Service Desk', receivedDateTime: inMinutes(-200), webLink: 'https://outlook.office.com/mail' }
    ],
    news: [
      { id: 'mock-news-1', title: 'New benefits enrollment window opens', webUrl: 'https://contoso.sharepoint.com/news/1', publishedDateTime: inMinutes(-1440) },
      { id: 'mock-news-2', title: 'Q3 all-hands recording now available', webUrl: 'https://contoso.sharepoint.com/news/2', publishedDateTime: inMinutes(-2880) },
      { id: 'mock-news-3', title: 'New collaboration space now open', webUrl: 'https://contoso.sharepoint.com/news/3', publishedDateTime: inMinutes(-4320) },
      { id: 'mock-news-4', title: 'Security awareness training due this month', webUrl: 'https://contoso.sharepoint.com/news/4', publishedDateTime: inMinutes(-5760) }
    ]
  };
}
