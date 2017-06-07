import {IToDoItem} from '../../model/IToDoItem';
import ItemDeletionCallBack from '../../model/ItemDeletionCallBack';

export  interface IToDoItemProps {
    item : IToDoItem;
    onItemRemoved : ItemDeletionCallBack;
}