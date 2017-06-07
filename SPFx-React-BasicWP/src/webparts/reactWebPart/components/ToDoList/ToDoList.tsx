import * as React from 'react';
import {IToDoListProps} from './IToDoListProps';
import {IToDoItem} from '../../model/IToDoItem';
import ToDoItem from '../ToDoItem/ToDoItem';

export default class ToDoList extends React.Component<IToDoListProps, {}>{
    constructor(props : IToDoListProps){
        super(props);
    }
    
    onItemRemoved = (item : IToDoItem)=>{
        this.props.onItemRemoved(item);
    }

    public render(){ 
        console.log(this.props.items);      
        let listItems = this.props.items.map((item : IToDoItem ) => {
            return (
                <ToDoItem item={item} onItemRemoved={this.onItemRemoved}/>
                );
        });
        return(
            <div>
                <ul>
                    {listItems}
                </ul>
            </div>
        );
    }
}