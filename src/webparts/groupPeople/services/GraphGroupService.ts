import { MSGraphClient } from '@microsoft/sp-http';
import IGraphGroupService from "../models/IGraphGroupService";
import IADGroup from '../models/IADGroup';
import IADGroupPeople from '../models/IADGroupPeople';

/**
 * SharePoint Group Service
 * 
 * REST API
 * @implements IGraphGroupService
 * @class
 */
export default class GraphGroupService implements IGraphGroupService {

    private ctx: any;

    public mClient: MSGraphClient;

    constructor(ctx) {
        this.ctx = ctx;
    }
    
    public initialize(): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            try {
                if(!this.mClient) {
                    this.ctx.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
                        this.mClient = client;
                        resolve(true);
                    });
                } else {
                    resolve(true);
                }
            } catch (error) {
                reject(false);                
            }
        })
    }

    public getADGroups(): Promise<Array<IADGroup>> {
        return new Promise((resolve, reject) => {
            try {
                this.mClient
                .api('/groups')
                .filter(`securityEnabled eq true`)
                .get((error, response: any, rawResponse?: any) => {
                    resolve(response.value);
                });
            } catch (error) {
                reject(error);                
            }
        })
    }
    
    public getADGroupPeoples(id: string): Promise<Array<IADGroupPeople>> {
        return new Promise((resolve, reject) => {
            try {
                this.mClient
                .api(`/groups/${id}/members`)
                .get((error, response: any, rawResponse?: any) => {
                    resolve(response.value);
                });
            } catch (error) {
                reject(error);                
            }
        })
    }
}