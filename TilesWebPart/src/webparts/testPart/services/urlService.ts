
export class UrlService {
 
    static BuildRequestUrl(siteUrl: string, listTitle: string, columns:string ,numberOfItems :number): string {
        let url = siteUrl.concat(`/_api/web/Lists/GetByTitle('${listTitle}')/items?$top=${numberOfItems.toString()}&$select=${columns},ID`);
        return url;
    }

    static BuildListItemUrl(baseUrl: string, listTitle: string, itemId: any){
        let url = baseUrl.concat(`/Lists/${listTitle}/DispForm.aspx?ID=${itemId}`);
        return url;
    }
}