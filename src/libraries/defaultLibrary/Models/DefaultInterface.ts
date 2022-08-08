import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { GraphFI } from "@pnp/graph";
import { SPFI } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface DunamicLibInterface {
  DynamicSPLoad(context?: ApplicationCustomizerContext | WebPartContext | null): Promise<SPFI>;
  DynamicGraphLoad(context?: ApplicationCustomizerContext | WebPartContext | null): Promise<GraphFI>;
}