import {IToDoItem} from '../model/IToDoItem';

export default class MockData {
    private static mockItems : IToDoItem[] = [{Id : 1, Title : "MockTask 1"}, 
                                                {Id : 2, Title : "MockTask 2"},
                                                {Id : 3, Title : "MockTask 3"},
                                                {Id : 4, Title : "MockTask 4"},];
    public static preData():IToDoItem[]{
        return MockData.mockItems;
    }
}
