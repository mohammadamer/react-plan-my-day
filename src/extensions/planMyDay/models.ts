export interface IMeeting {
  id: string;
  subject: string;
  start: Date;
  end: Date;
  joinUrl?: string;
  isAllDay: boolean;
}

export interface IPlannerTask {
  id: string;
  title: string;
  dueDateTime?: Date;
  planTitle: string;
  percentComplete: number;
}

export interface IImportantEmail {
  id: string;
  subject: string;
  from: string;
  receivedDateTime: Date;
  webLink: string;
}

export interface INewsPost {
  id: string;
  title: string;
  webUrl: string;
  publishedDateTime: Date;
  imageUrl?: string;
}

export interface ICopilotAnswer {
  text: string;
}

export interface IPlanMyDayData {
  displayName: string;
  meetings: IMeeting[];
  tasks: IPlannerTask[];
  emails: IImportantEmail[];
  news: INewsPost[];
}
