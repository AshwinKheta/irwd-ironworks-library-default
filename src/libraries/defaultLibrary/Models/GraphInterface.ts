export interface GraphInterface {
  getSiteId(siteUrl?: string): Promise<string>;
  getListId(listName: string, siteUrl?: string): Promise<string>;
  getItems(concatString: string, select?: string[], filter?: string, orderBy?: any[], top?: number): Promise<any | null>;
  getTargetFilter(): Promise<string>;
  formatItems(result: any, selectedFields: string[], requiredFields?: string[], optionalFields?: string[], modifiedfields?: string[]): any[];
}