import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ClientMode } from "./ClientMode";

export interface IGraphProps {
  clientMode: ClientMode;
  context: WebPartContext;
}
