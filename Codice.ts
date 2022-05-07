// NamedRanges
const NR_COPPIE = "COPPIE"; // Coppia OID - AULA
const NR_OID = "OID";
const NR_AULE = "AULE";
const NR_LUNEDI = "LUNEDI";
const NR_MARTEDI = "MARTEDI";
const NR_MERCOLEDI = "MERCOLEDI";
const NR_GIOVEDI = "GIOVEDI";
const NR_VENERDI = "VENERDI";

interface SingleEvent {
    title: string,
    start: Date,
    end: Date,
    allDay: boolean
}

interface Date {
    GetFirstDayOfWeek(): Date,
    GetLastDayOfWeek(): Date,
}

Date.prototype.GetFirstDayOfWeek = function() {
    return (new Date(this.setDate(this.getDate() - this.getDay()+ (this.getDay() == 0 ? -6:1) )));
}
Date.prototype.GetLastDayOfWeek = function() {
    return (new Date(this.setDate(this.getDate() - this.getDay() +7)));
}

var oids: Array<string> = [];
var aule: Array<string> = [];
var coppie: {[oid: string]: {aula: string, row: number, events: Array<SingleEvent>}} = {};

function fetchListaAule(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let result = ss.getRangeByName(NR_COPPIE).getValues();
    result.shift();
    result.forEach((c, i) => {
        if (c[0] === "" || c[1] === "") return
        coppie[c[0]] = {aula: c[1], row: i+2, events: null};
        oids.push(c[0]);
        aule.push(c[1]);
    })
}

function parseSingleResponse(response: GoogleAppsScript.URL_Fetch.HTTPResponse, oid: string = null){
    if (oid === null){
        const headers: any = response.getHeaders();
        oid = headers.arguments[0];
    }

    if (response.getResponseCode() === 200){
        let text = response.getContentText();
        const pattern = /(var events = \[.*?\];)/sg;
        const found = text.match(pattern)[0];
        var events: Array<SingleEvent>;
        eval(found);
        coppie[oid].events = filter_week(events.slice(), new Date());
    }
    else {
        Logger.log("pobblema");
    }
}

function getSingleCalendar(oid: string){
    const url = `https://offweb.unipa.it/offweb/public/aula/calendar.seam?oidAula=${oid}`;
    let response = UrlFetchApp.fetch(url);
    parseSingleResponse(response);
}


function mainFunction() {
    fetchListaAule();
    const urls = oids.map(oid => `https://offweb.unipa.it/offweb/public/aula/calendar.seam?oidAula=${oid}`);
    let responses: Array<GoogleAppsScript.URL_Fetch.HTTPResponse> = [];
    const request_limit = 100 - (1);
    for (let i = 0; i < urls.length; i += request_limit) {
        const batch = urls.slice(i, i + request_limit);
        Logger.log(`Sending ${batch.length} requests...`);
        responses = responses.concat(UrlFetchApp.fetchAll(batch));
        Logger.log("Sent")
    }
    responses.forEach((resp, i) => parseSingleResponse(resp, oids[i]));
}



function filter_week(events: Array<SingleEvent>, day: Date): Array<SingleEvent>{
    let first = day.GetFirstDayOfWeek();
    let last = day.GetLastDayOfWeek();
    let filtered = events.filter(e => first < e.start && e.start < last && e.start.getFullYear() == first.getFullYear());
    // Logger.log(`Partendo da ${events.length} to ${filtered.length}`)
    return filtered.slice()
}