import "@pnp/core";
import "@pnp/graph";
import "@pnp/graph/calendars";
import "@pnp/graph/contacts";
import "@pnp/graph/graphqueryable";
import "@pnp/graph/groups";
import "@pnp/graph/planner";
import "@pnp/graph/sites";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import "@pnp/logging";
import "@pnp/queryable";

import { Caching, InjectHeaders } from "@pnp/queryable";
import { GraphFI, SPFx, graphfi } from "@pnp/graph";
import { GraphInterface, SPInterface } from "../Models";
import { LogLevel, PnPLogging } from "@pnp/logging";

import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class GRAPH implements GraphInterface {
  private context: ApplicationCustomizerContext | WebPartContext;
  public g: GraphFI = null;
  public sp: SPInterface;
  public constructor(Icontext: ApplicationCustomizerContext | WebPartContext) {
    this.context = Icontext;
    let o = { 'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly' };
    if (this.g == null)
      this.g = graphfi().using(SPFx(this.context)).using(InjectHeaders(o)).using(Caching()).using(PnPLogging(LogLevel.Verbose));
  }
  public async importDynamicSubDep() {
    let s = await import("./SP");
    return this.sp = await new s.default(this.context);
  }
  public async getSiteId(siteUrl?: string) {
    return await this.sp.getSiteId(siteUrl);
  }
  public async getListId(listName: string, siteUrl?: string) {
    return await this.sp.getListId(listName, siteUrl);
  }
  public async getItems(concatString: string, select?: string[], filter?: string, orderBy?: any[], top?: number) {
    let res = [];
    if (select && select.length) {
      concatString += "($select=";
      for (let i = 0; i < select.length; i++) {
        const s = select[i];
        if (i < select.length - 1)
          concatString += "" + s.trim() + ",";
        else
          concatString += "" + s.trim();
      }
      concatString += ")";
    }
    if (filter && filter.length) {
      if (filter.indexOf("fields/") > -1) {
        concatString += "&$filter=" + filter;
      }
    }
    if (orderBy && orderBy.length) {
      concatString += "&$orderBy=";
      let s = [];
      for (let i = 0; i < orderBy.length; i++) {
        const o = orderBy[i];
        s.push("" + o.name + ("" + (o.order ? "" : " desc")));
      }
      concatString += s.join(",");
    }
    if (top && top > 0) {
      concatString += "&$top=" + top
    }
    res = await this.g.sites.concat(concatString)()
      .then((e: any) => {
        if (e && e.length) {
          return e;
        } else return [];
      })
      .catch(er => (console.log(er), console.log(er.message), []));
  }
  public formatItems(result: any, selectedFields: string[], requiredFields?: string[], optionalFields?: string[], modifiedfields?: string[]) {
    if (result && result.length) {
      let v: any[] = [];
      for (let im = 0; im < result.length; im++) {
        const f = result[im];
        let t: any = {};
        if (f && f.fields) {
          let ff = f.fields;
          if (optionalFields && optionalFields.length) {
            for (let io = 0; io < optionalFields.length; io++) {
              const o = optionalFields[io];
              t[o] = ff[o] ? ff[o] : null;
            }
          }
          if (selectedFields && selectedFields.length)
            for (let iis = 0; iis < selectedFields.length; iis++) {
              const e = selectedFields[iis];
              t[e] = ff[e];
            }
          if (requiredFields && requiredFields.length)
            for (let ir = 0; ir < requiredFields.length; ir++) {
              const r = requiredFields[ir];
              if (ff[r] && ff[r].length) { } else {
                t = null;
              }
            }
          if (modifiedfields && modifiedfields.length)
            for (let im = 0; im < modifiedfields.length; im++) {
              const m = modifiedfields[im];
              if (m && m.length && m.split(":") && m.split(":").length == 2) {
                let origFN = m.split(":")[0];
                let newPN = m.split(":")[1];
                t[newPN] = ff[origFN];
              }
            }
          v.push(t)
        }
      }
      v.filter(l => l != null);
      return v;
    } else return [];
  }
  public async getTargetFilter() {
    return await this.sp.getTargetFilter(true);
  }
}