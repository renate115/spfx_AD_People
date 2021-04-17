import { MSGraphClient } from '@microsoft/sp-http';
import IADGroup from './IADGroup';
import IADGroupPeople from './IADGroupPeople';

export default interface IGraphGroupService {
    mClient: MSGraphClient;
    initialize(): Promise<boolean>;
    getADGroups(): Promise<Array<IADGroup>>;
    getADGroupPeoples(id: string): Promise<Array<IADGroupPeople>>;
}