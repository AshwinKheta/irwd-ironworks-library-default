import "@pnp/core"
import "@pnp/logging"
import "@pnp/queryable"
import "@pnp/sp"
import "@pnp/sp/attachments"
import "@pnp/sp/fields"
import "@pnp/sp/files"
import "@pnp/sp/folders"
import "@pnp/sp/items"
import "@pnp/sp/lists"
import "@pnp/sp/profiles"
import "@pnp/sp/sites"
import "@pnp/sp/webparts"
import "@pnp/sp/webs"

import { ISPProps, SPInterface } from "./../Models";
import { ISite, Site } from "@pnp/sp/sites"
import { IWeb, Web } from "@pnp/sp/webs"
import { LogLevel, PnPLogging } from "@pnp/logging"
import { SPFI, SPFx, spfi } from "@pnp/sp"

import { ApplicationCustomizerContext } from "@microsoft/sp-application-base"
import { BaseWebPartContext } from "@microsoft/sp-webpart-base"
import { Caching } from "@pnp/queryable"

export default class SP implements SPInterface {
  private context: ApplicationCustomizerContext | BaseWebPartContext;
  public sp: SPFI = null;
  public constructor(Icontext: ApplicationCustomizerContext | BaseWebPartContext) {
    this.context = Icontext;
    if (this.sp == null)
      this.sp = spfi().using(SPFx(this.context)).using(Caching()).using(PnPLogging(LogLevel.Verbose));
  }
  private async checkSiteUrl(siteUrl?: string): Promise<boolean> {
    if (siteUrl && siteUrl.length) {
      let web = await this.sp.web.select("Url")();
      return siteUrl && siteUrl.length && siteUrl.toLowerCase() != web.Url.toLowerCase()
    } else return Promise.resolve(false);
  }
  public async getSiteId(siteUrl?: string): Promise<string> {
    let siteId = "";
    siteId = await this.sp.site.select("Id")().then(s => s.Id).catch(ex => (console.log(ex), ""));
    if (await this.checkSiteUrl(siteUrl)) {
      let site: ISite = Site(siteUrl).using(SPFx(this.context));
      siteId = await site.select("Id")().then(s => s.Id).catch(er => (console.log(er), ""));
    }
    return siteId;
  }
  public async getListId(listName: string, siteUrl?: string): Promise<string> {
    let lists = (await this.checkSiteUrl(siteUrl))
      ? Web(siteUrl).using(SPFx(this.context)).lists
      : this.sp.web.lists;
    return await lists.getByTitle(listName)().then(l => l.Id).catch(er => (console.log(er), ""));
  }
  public async getItemsByListName(listName: string, select?: string[], filter?: string, orderBy?: any[], expand?: string[], top?: number, skip?: number) {
    let list = this.sp.web.lists.getByTitle(listName);
    return await this.getItems(list, select, filter, orderBy, expand, top, skip);
  }
  public async getItemsByListId(listId: string, select?: string[], filter?: string, orderBy?: any[], expand?: string[], top?: number, skip?: number) {
    let list = this.sp.web.lists.getById(listId);
    return await this.getItems(list, select, filter, orderBy, expand, top, skip);
  }
  public async getTargetFilter(forGraph?: boolean) {
    return await this.sp.profiles.myProperties
      .select("UserProfileProperties")()
      .then((userProperties: any) => {
        let trueValuedUserProfileProperties: string[] = [],
          requiredUserProfilePropertyKeys: string[] = [
            "IsHeadquarterContractor",
            "IsHeadquarterEmployee",
            "IsColleague",
            "IsFieldSalesForceContractor",
            "IsFieldSalesForceEmployee",
            "IsLabWorker"
          ],
          relativeTargetOptions: string[] = [
            "All Headquarter Contractors",
            "All Headquarter Employees",
            "All Employees",
            "All Field Sales Contractors",
            "All Field Sales Employees",
            "All Lab Employees & Contractors",
            "All Employees & Contractors"
          ],
          filteredTargetOptionsfilterStrings: string[] = [],
          filetredUserProfileProperties: any[] = userProperties && userProperties.UserProfileProperties && userProperties.UserProfileProperties.length ? userProperties.UserProfileProperties.filter((userProfileProperty: any) => requiredUserProfilePropertyKeys.indexOf(userProfileProperty.Key) > -1) : [];
        filetredUserProfileProperties.map((userProfileProperty: any) => {
          userProfileProperty.Value == "True" ?
            trueValuedUserProfileProperties.push(userProfileProperty.Key)
            : console.log("not true for " + userProfileProperty.Key);
        });
        if (trueValuedUserProfileProperties && trueValuedUserProfileProperties.length) {
          trueValuedUserProfileProperties.map((trueValuedUserProfileProperty: string) => {
            if (requiredUserProfilePropertyKeys.indexOf(trueValuedUserProfileProperty) > -1)
              filteredTargetOptionsfilterStrings.push("" + (forGraph ? "fields/" : "") + "IWTarget eq '" + encodeURIComponent(relativeTargetOptions[requiredUserProfilePropertyKeys.indexOf(trueValuedUserProfileProperty)]) + "'");
          });
        }
        filteredTargetOptionsfilterStrings.push("" + (forGraph ? "fields/" : "") + "IWTarget eq '" + encodeURIComponent(relativeTargetOptions[6]) + "'");
        return "(" + filteredTargetOptionsfilterStrings.join(" or ") + ")";
      })
      .catch((errorFetchingUserProfileProperties) => {
        console.log(errorFetchingUserProfileProperties);
        return "";
      });
  }
  private async getItems(list: any, select?: string[], filter?: string, orderBy?: any[], expand?: string[], top?: number, skip?: number) {
    let items = list.items;
    if (select && select.length)
      items.select(...select);
    if (expand && expand.length)
      items.expand(...expand);
    if (filter && filter.length)
      items.filter(filter);
    if (orderBy && orderBy.length)
      for (let i = 0; i < orderBy.length; i++) {
        const o = orderBy[i];
        items.orderBy(o.name, o.order);
      }
    if (top != null)
      items.top(top);
    if (skip != null)
      items.skip(skip);
    return await items();
  }
}