import { IUserItem } from "./IUserItem";

export interface IGraphState {
  users: Array<IUserItem>;
  searchFor: string;
}
