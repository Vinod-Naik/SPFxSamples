import * as React from 'react';
import { IToDoContainerProps } from './IToDoContainerProps';
import { IToDoContainerState } from './IToDoContainerState';
import ToDoList from '../ToDoList/ToDoList';
import ToDoForm from '../ToDoForm/ToDoForm';
import * as MockData from '../../dataprovider/mockdata';
import {IToDoItem} from '../../model/IToDoItem';

export default class ToDoContainer extends React.Component<IToDoContainerProps , IToDoContainerState> {
    constructor(props : IToDoContainerProps){
        super(props);
        this.state = {
            todoItems : this.props.items
        };
    }

    onItemAdded = (inputValue)=>{
        console.log("Item added called");
        let updatedTasks : IToDoItem[] = this.props.items;
        let lastItemId = updatedTasks[updatedTasks.length -1].Id
        updatedTasks.push({Id : lastItemId + 1, Title : inputValue});
        this.setState({ 
            todoItems : updatedTasks
        });
    }
    
    onItemRemoved = (item : IToDoItem)=>{
        console.log("Task removed");
        let updatedTasks : IToDoItem[] = this.props.items;
        let i = updatedTasks.indexOf(item);
        updatedTasks.splice(i, 1);
        //let index = updatedTasks.findIndex(item => item.key == itemKey);
        this.setState({ todoItems : updatedTasks });
    }

    public render() {
        return (
            <div>
                <span>Welcome to SharePoint!</span>
                <span>WebPart Description : {this.props.description}</span>
                <div>
                    <ToDoForm onItemAdded={this.onItemAdded}/>
                    <ToDoList items={this.state.todoItems} onItemRemoved={this.onItemRemoved}/>
                </div>
            </div>
        );
    }
}