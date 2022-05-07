// NamedRanges
const NR_COPPIE = "COPPIE"; // Coppia OID - AULA
const NR_OID = "OID";
const NR_AULE = "AULE";
const NR_LUNEDI = "LUNEDI";
const NR_MARTEDI = "MARTEDI";
const NR_MERCOLEDI = "MERCOLEDI";
const NR_GIOVEDI = "GIOVEDI";
const NR_VENERDI = "VENERDI";
const FREE = "FREE";
const OCC = "OCC";
const CACHE_DUR = 60 * 60 * 12; // 12h

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

function packCoppie() :{[oid: string]: string}{
    const cache = CacheService.getDocumentCache()
    var copia: {[oid: string]: string} = {}
    oids.forEach(oid => {
        copia[oid] = JSON.stringify(coppie[oid]);
    })
    cache.putAll(copia);
    return copia
}

function unpackCoppie(){
    const cache = CacheService.getDocumentCache();
    const datas = cache.getAll(oids);
    if(Object.keys(datas).length !== oids.length) return null;
    Object.entries(datas).map((entry) => {
        let key = entry[0];
        let value = entry[1];
        coppie[key] = JSON.parse(value);
        coppie[key].events.forEach((ev, index) => {
            coppie[key].events[index].start = new Date(ev.start);
            coppie[key].events[index].end = new Date(ev.end);
        })
    })
    return coppie
}

function fetchListaAule(ss: GoogleAppsScript.Spreadsheet.Spreadsheet = null){
    let cache = CacheService.getDocumentCache();
    let temp_oids = cache.get("oids")
    let temp_aule = cache.get("aule");
    if (temp_aule !== null && temp_oids !== null){
        aule = JSON.parse(temp_aule);
        oids = JSON.parse(temp_oids);
        oids.forEach((oid, i) => {
            coppie[oid] = {aula: aule[i], row: i+2, events: null}
        })
        return;
    }
    if (ss === null)
        ss = SpreadsheetApp.getActiveSpreadsheet();
    let result = ss.getRangeByName(NR_COPPIE).getValues();
    result.shift();
    result.forEach((c, i) => {
        if (c[0] === "" || c[1] === "") return
        coppie[c[0]] = {aula: c[1], row: i+2, events: null};
        oids.push(c[0]);
        aule.push(c[1]);
    cache.put("aule", JSON.stringify(aule), CACHE_DUR);
    cache.put("oids", JSON.stringify(oids), CACHE_DUR);
}

function filter_week(events: Array<SingleEvent>, day: Date): Array<SingleEvent>{
    let first = day.GetFirstDayOfWeek();
    let last = day.GetLastDayOfWeek();
    first.setHours(0,0,0);
    last.setHours(23,59,59);
    let filtered = events.filter(e => first <= e.start && e.start <= last && e.start.getFullYear() == first.getFullYear());
    // Logger.log(`Partendo da ${events.length} to ${filtered.length}`)
    return filtered.slice()
}

function parseSingleResponse(response: GoogleAppsScript.URL_Fetch.HTTPResponse, oid: string = null, day: Date = null){
    if (oid === null){
        const headers: any = response.getHeaders();
        oid = headers.arguments[0];
    }
    if (day === null){
        day = new Date();
    }

    if (response.getResponseCode() === 200){
        let text = response.getContentText();
        const pattern = /(var events = \[.*?\];)/sg;
        const found = text.match(pattern)[0];
        var events: Array<SingleEvent>;
        eval(found);
        coppie[oid].events = filter_week(events.slice(), day);
    }
    else {
        Logger.log("pobblema");
    }
}

function updateView(ss: GoogleAppsScript.Spreadsheet.Spreadsheet = null){
    if (ss === null)
        ss = SpreadsheetApp.getActiveSpreadsheet();
    const sss = ss.getActiveSheet() || ss.getSheets()[0];
    const days = [NR_LUNEDI, NR_MARTEDI, NR_MERCOLEDI, NR_GIOVEDI, NR_VENERDI]
    days.forEach((ns_day, day_index) => {
        const day_range = ss.getRangeByName(ns_day)
        day_range.offset(1,0).setValue(FREE);
        const first_column = day_range.getColumn();
        const first_hour = 8;
        oids.forEach(oid => {
            const u = coppie[oid];
            u.events.forEach( event => {
                if (event.start.getUTCDay() != (day_index + 1)) return
                const duration = event.end.getHours() - event.start.getHours();
                const rangeToEdit = sss.getRange(u.row, first_column + (event.start.getHours() - first_hour), 1, duration);
                rangeToEdit.setValue(OCC)
                Logger.log(`OID: ${oid}, Aula: ${u.aula},D:${event.start.getDay()}, ${event.start.getHours()}:${event.start.getMinutes()}-${event.end.getHours()}:${event.end.getMinutes()}, ${rangeToEdit.getA1Notation()}`)
            })
        })
        day_range.offset(Math.max(oids.length, aule.length) + 1, 0).clearContent();
    })
}

function fetchAllData(){
    fetchListaAule();
    const cache = CacheService.getDocumentCache();
    let temp = unpackCoppie()
    if (temp !== null){
        Logger.log("Utilizzando i dati cachati")
        return coppie;
    }
    Logger.log("Nessun dato cachato...")
    const urls = oids.map(oid => `https://offweb.unipa.it/offweb/public/aula/calendar.seam?oidAula=${oid}`);
    let responses: Array<GoogleAppsScript.URL_Fetch.HTTPResponse> = [];
    const request_limit = 100 - (1);
    for (let i = 0; i < urls.length; i += request_limit) {
        const batch = urls.slice(i, i + request_limit);
        Logger.log(`Sending ${batch.length} requests...`);
        responses = responses.concat(UrlFetchApp.fetchAll(batch));
        Logger.log("Sent")
    }
    let day = new Date();
    if (false){
        day = new Date(day.setDate(day.getDate() + 7))
    }
    responses.forEach((resp, i) => parseSingleResponse(resp, oids[i], day));
    packCoppie()
    return coppie;
}

function mainFunction(e: GoogleAppsScript.Events.TimeDriven) {
    fetchAllData();
    updateView();
}
