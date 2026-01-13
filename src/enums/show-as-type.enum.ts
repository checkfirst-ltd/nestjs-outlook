/**
 * FreeBusyStatus enum representing the availability status of a calendar event
 * Matches Microsoft Graph FreeBusyStatus type
 *
 * @see https://learn.microsoft.com/en-us/graph/api/resources/event
 */
export enum ShowAsType {
  UNKNOWN = "unknown",
  FREE = "free",
  TENTATIVE = "tentative",
  BUSY = "busy",
  OOF = "oof",
  WORKING_ELSEWHERE = "workingElsewhere",
}
