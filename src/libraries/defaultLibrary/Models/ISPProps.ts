import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";

export interface ISPProps {
  context: ApplicationCustomizerContext | BaseWebPartContext;
}