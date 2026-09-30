export function isMockMode(): boolean {
  return new URLSearchParams(window.location.search).get('mockData') === 'true';
}

export function getGreetingKey(): 'GreetingMorning' | 'GreetingAfternoon' | 'GreetingEvening' {
  const hour: number = new Date().getHours();
  if (hour < 12) {
    return 'GreetingMorning';
  }
  if (hour < 18) {
    return 'GreetingAfternoon';
  }
  return 'GreetingEvening';
}

export function formatTime(date: Date): string {
  return date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}

export function formatTemplate(template: string, ...args: string[]): string {
  return template.replace(/{(\d+)}/g, (_match, index) => args[Number(index)] ?? '');
}
