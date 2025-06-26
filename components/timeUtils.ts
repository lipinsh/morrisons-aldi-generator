export const subtractHour = (timeStr: string): string => {
  const [hours, minutes] = timeStr.split(':').map(Number);
  const newHours = (hours - 1 + 24) % 24;
  return `${String(newHours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
};

export const timeToMinutes = (timeStr: string): number => {
  const [hours, minutes] = timeStr.split(':').map(Number);
  return hours * 60 + minutes;
};

export const minutesToTime = (minutes: number): string => {
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
};
