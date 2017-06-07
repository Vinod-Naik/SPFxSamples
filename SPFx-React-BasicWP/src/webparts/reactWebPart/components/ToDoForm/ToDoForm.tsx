import * as React from 'react';
import {IToDoFormProps} from './IToDoFormProps';

export default class ToDoForm extends React.Component<IToDoFormProps, {}>{
    constructor(props: IToDoFormProps) {
        super(props);
    }
    private onFormSubmit = (e)=>{
        e.preventDefault();
        var ipElm = e.target.querySelector('input');
        if(ipElm.value != ''){
            console.log('Form Submitted');
            this.props.onItemAdded(ipElm.value);
        }
        ipElm.value = '';
    }
    public render(){
        return(
            <div>
                <form onSubmit={this.onFormSubmit}>
                    <input type='text'></input>
                </form>
            </div>
        );
    }
}