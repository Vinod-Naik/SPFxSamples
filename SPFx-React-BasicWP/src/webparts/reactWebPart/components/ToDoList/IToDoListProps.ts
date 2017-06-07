import {IToDoItem} from '../../model/IToDoItem';
import ItemDeletionCallBack from '../../model/ItemDeletionCallBack';

export interface IToDoListProps {
    items : IToDoItem[];
    onItemRemoved : ItemDeletionCallBack;
}