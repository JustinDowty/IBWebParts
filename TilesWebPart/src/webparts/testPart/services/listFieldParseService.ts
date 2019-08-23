
export class ListFieldParseService {

    public ExtractFieldValue(fieldObject: any, TaxCatchAll: any): string {

        var value = '';
        if (Array.isArray(fieldObject)) {
            $(fieldObject).each(function (index, item) {
                var termId = item.WssId;
                var term = TaxCatchAll.find(function(item:any){ return item.ID === termId;});
                value += `${term.Term}, `;
            });

            value = value.replace(/,\s*$/, "");
            return value;
            
        } else if (fieldObject && fieldObject.hasOwnProperty('Label') 
                && fieldObject.hasOwnProperty('TermGuid')) {
                    var termId = fieldObject.WssId;
                    var term = TaxCatchAll.find(function(item:any){ return item.ID === termId;});
                    value = term.Term;
                    return value;
        }

        value = fieldObject;
        return value;
    }
}