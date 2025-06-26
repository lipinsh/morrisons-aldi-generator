import { describe, it, expect } from 'vitest';
import { subtractHour, timeToMinutes, minutesToTime } from '../timeUtils';

describe('time utils', () => {
  it("subtractHour wraps around midnight", () => {
    expect(subtractHour('00:30')).toBe('23:30');
  });

  it('timeToMinutes converts correctly', () => {
    expect(timeToMinutes('02:15')).toBe(135);
  });

  it('minutesToTime converts correctly', () => {
    expect(minutesToTime(135)).toBe('02:15');
  });
});
