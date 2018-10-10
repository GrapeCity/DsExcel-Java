import { Component } from '@angular/core';
@Component({
    selector: 'home',
    templateUrl: './home.component.html'
})

export class HomeComponent {
    options: Array<object> = [
        { value: 'BidTracker', label: 'Bid Tracker' },
        { value: 'ToDoList', label: 'ToDo List' },
        { value: 'AddressBook', label: 'Address Book' }
    ];
    selectedValue: string;

    constructor() {
        this.selectedValue = "BidTracker";
    }

    onChange(newValue) {
        console.log(newValue);
        this.selectedValue = newValue;
        // ... do other stuff here ...
    }
}
