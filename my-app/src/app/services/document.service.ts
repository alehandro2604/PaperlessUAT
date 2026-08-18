import { Injectable } from '@angular/core';
import { SharePointListService } from './sharepoint-list.service';

@Injectable({
  providedIn: 'root'
})
export class DocumentService {
  constructor(
    private listService: SharePointListService
  ) { }

  async getLibraries() {
    const lists = await this.listService.getAllLists();

    return lists.filter(
      (x: any) => x.template === 'documentLibrary'
    );
  }
}
