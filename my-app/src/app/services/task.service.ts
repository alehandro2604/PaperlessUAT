import { Injectable } from '@angular/core';
import { SharePointListService } from './sharepoint-list.service';


@Injectable({
    providedIn: 'root'
})
export class TaskService {
    constructor(
        private listService: SharePointListService
    ) { }

    async loadTaskLists() {
        const lists = await this.listService.getAllLists();

        return lists.filter(
            (list: any) =>
                list.displayName
                    .toLowerCase()
                    .includes('task')
        ).map((list: any) => list.displayName ?? list.name);
    }

    async loadTasks(listName: string) {
        const list = await this.listService.getListByName(listName);
        if (!list)
            return [];
        return await this.listService.getListItems(list.id, list.parentReference?.siteId);
    }
}
