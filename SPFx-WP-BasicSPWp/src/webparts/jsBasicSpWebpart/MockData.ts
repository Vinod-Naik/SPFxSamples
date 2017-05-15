
import { ISPList } from './JsBasicSpWebpartWebPart';

export default class MockData{
    private static _items :ISPList[] = [{Title : "List 1",  Id:"1"},
                                        {Title : "List 2",  Id:"2"},
                                        {Title : "List 3",  Id:"3"}    ];

    public static get(restUrl : string, options?:any):Promise<ISPList[]>{
        return new Promise<ISPList[]>((resolve)=>{
            resolve(MockData._items);
        });     
    }
}