export interface SPInterface {
  getSiteId(siteUrl?: string): Promise<string>;
  getListId(listName: string, siteUrl?: string): Promise<string>;
  getItemsByListName(listName: string, select?: string[], filter?: string, orderBy?: any[], expand?: string[], top?: number, skip?: number): Promise<any[] | void>;
  getItemsByListId(listId: string, select?: string[], filter?: string, orderBy?: any[], expand?: string[], top?: number, skip?: number): Promise<any[] | void>;
  getTargetFilter(forGraph?: boolean): Promise<string>;
}