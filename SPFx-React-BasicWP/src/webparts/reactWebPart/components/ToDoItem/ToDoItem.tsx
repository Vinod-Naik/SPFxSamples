import * as React from 'react';
import { IToDoItemProps } from './IToDoItemProps';

export default class IToDoItem extends React.Component<IToDoItemProps, {}> {
    constructor(props){
        super(props);
    }

    removeTask = (e)=>{
        e.preventDefault();
        this.props.onItemRemoved(this.props.item);
    }


    public render() {
        return (
            <li key={this.props.item.Id}>
                    <span>{this.props.item.Title} </span>
                    <button onClick={this.removeTask}>X</button>
            </li>
        );
    }
}