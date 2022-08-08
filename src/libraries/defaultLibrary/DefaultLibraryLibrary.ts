import { GraphInterface, SPInterface } from "./Models";

import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export * from "./Models";
export class DefaultLibraryLibrary {
  public name(): string {
    return 'DefaultLibraryLibrary';
  }
  public InterfaceLoad() {
    return import("./Models");
  }
  public async DynamicSPLoad(context?: ApplicationCustomizerContext | WebPartContext): Promise<any> {
    const spPnP = await import("./SubLibraries/SP");
    return (await new spPnP.default(context)) as SPInterface;
  }
  public async DynamicGraphLoad(context?: ApplicationCustomizerContext | WebPartContext): Promise<any> {
    const g = await import("./SubLibraries/Graph");
    return (await new g.default(context) as GraphInterface);
  }
}