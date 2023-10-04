var expressions = require('angular-expressions');
var translate = require('../translate')
var html2ooxml = require('./html2ooxml');

var numbers = {'nl': ['nul', 'een', 'twee', 'drie', 'vier', 'vijf', 'zes', 'zeven', 
                      'acht', 'negen', 'tien', 'elf', 'twaalf', 'dertien', 'veertien', 
                      'vijftien', 'zestien', 'zeventien', 'achtien', 'negentien', 'twintig']};

// Apply all customs functions
function apply(data) {

}
exports.apply = apply;

// *** Custom modifications of audit data for usage in word template


// *** Custome Angular expressions filters ***

var filters = {};

expressions.filters.cveSlider = function(input, score) {
    score = typeof score == "String" ? parseInt(score) : score;
    if (score == undefined || score < 0 || score > 10)
        score = 0;

    return "/app/images/sliders/Slider_groen-rood_" + Math.round(score)*10 + ".png";
}

expressions.filters.cveNumber = function(input, score) {
    return `${' '.repeat(Math.round(score)*6)}${score}`;
}

expressions.filters.check = function(input, check) {
    return (input && input.some((x) => x == check));
}

expressions.filters.condCheck = function(input) {
    return `<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /><w:jc w:val="center" /></w:pPr><w:r>${(input) ? '<w:sym w:font="Wingdings" w:char="F0FC" />' : ''}</w:r></w:p>`;
}

expressions.filters.extractTargets = function(input) {
    return [...new Set(input.match(/((https?:\/\/){0,1}[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*))/g))] ?? [];
}

expressions.filters.generateTargetsTable = function(input) {
    tw = 5100
    cw = tw / 6
    
    // https://docxperiments.readthedocs.io/en/latest/synthesis/documentxml.html
    pre = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid" /><w:tblW w:w="${tw}" w:type="pct" /><w:tblBorders><w:top w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:left w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:bottom w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:right w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideH w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideV w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /></w:tblBorders><w:tblLayout w:type="fixed" /></w:tblPr><w:tblGrid><w:gridCol w:w="${cw}" /><w:gridCol w:w="${cw}" /><w:gridCol w:w="${cw}" /><w:gridCol w:w="${cw}" /><w:gridCol w:w="${cw}" /><w:gridCol w:w="${cw}" /></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${tw}" w:type="dxa" /><w:gridSpan w:val="6" /><w:shd w:val="clear" w:color="auto" w:fill="2DA5DE" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF" w:themeColor="background1" /></w:rPr><w:t xml:space="preserve">Doelobjecten</w:t></w:r></w:p></w:tc></w:tr>`;
    post = '</w:tbl>';

    // Split array into equally sized portions
    var rows = []
    while (input.length > 0)
        rows.push(input.splice(0, 6));
    
    out = "";
    for (row of rows) {
        out += "<w:tr>";
        l = (row.length == 6) ? false : row.length-1;
        for ([i, t] of row.entries()) {
            if (l && i == l)
                out += `<w:tc><w:tcPr><w:tcW w:w="${cw*(6-l)}" w:type="dxa" /><w:gridSpan w:val="${6-l}" /><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1" w:themeFillShade="D9" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
            else
                out += `<w:tc><w:tcPr><w:tcW w:w="${cw}" w:type="dxa" /><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1" w:themeFillShade="D9" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
        }
        out += "</w:tr>"
    }

    return pre + out + post;
}

expressions.filters.extractTable = function(input) {
    if (input == null)
        return [];
    table = input[0].text.match(/<table>(.*)<\/table>/gm);
    if (table == null)
        return [];
    table = table[0];
    rows = table.match(/<tr>(.*?)<\/tr>/gm);
    if (rows == null)
        return [];
    // Somehow, without the double map (first using regular 'match' instead of 'matchAll'),
    // the whole string got matched, not just the capture group.
    return rows.map((row) => [...row.matchAll(/<td.*?>(.*?)<\/td>/gm)].map((match) => match[1]));
}

expressions.filters.generateOutdatedSoftwareTable = function(input) {
    if (input == null || input.length == 0)
        return '';
    cols = input[0].length;
    switch (cols) {
        case 2:
            var tw = 4000;
            var col_width = [1200, 2800];
            var col_names = ['Doelobjecten', 'Geïnstalleerde versie/Oplossing'];
            break;
        case 3:
            var tw = 5100;
            var col_width = [1200, 2100, 1800];
            var col_names = ['Kwetsbaar product', 'Geïnstalleerde versie/Oplossing', 'Doelobjecten'];
            break;
        case 4:
            var tw = 5100;
            var col_width = [1200, 1800, 600, 1500];
            var col_names = ['Kwetsbaar product', 'Geïnstalleerde versie/Oplossing', 'CVSS', 'Doelobjecten'];
            break;
        default:
            console.log(`Outdated Software table has invalid number of columns (${cols})! not parsing. Data: ${input}`);
            return "";
    }
    
    // https://docxperiments.readthedocs.io/en/latest/synthesis/documentxml.html
    pre = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid" /><w:tblW w:w="${tw}" w:type="pct" /><w:tblBorders><w:top w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:left w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:bottom w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:right w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideH w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideV w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /></w:tblBorders><w:tblLayout w:type="fixed" /></w:tblPr>`;
    pre += '<w:tblGrid>' + col_width.map((w) => ` <w:gridCol w:w="${w}" />`).join('') + '</w:tblGrid>';
    pre += '<w:tr>' + col_width.map((w, idx) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa" /><w:shd w:val="clear" w:color="auto" w:fill="2DA5DE" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF" w:themeColor="background1" /></w:rPr><w:t xml:space="preserve">${col_names[idx]}</w:t></w:r></w:p></w:tc>`).join('') + '</w:tr>' 
    post = '</w:tbl>';
    
    out = "";
    for (row of input) {
        out += "<w:tr>";
        for (var i = 0; i < cols; i++) {
            data = html2ooxml(row[i].replace(/(<p><\/p>)+$/, ''))
            out += `<w:tc><w:tcPr><w:tcW w:w="${col_width[i]}" w:type="dxa" /><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1" w:themeFillShade="D9" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:t>${data}</w:t></w:r></w:p></w:tc>`
        }
        out += "</w:tr>";
    }

    return pre + out + post;
}

function count_findings_per_severity(findings, severity) {
    return findings.filter((finding) => finding.cvss['environmentalSeverity'] == severity).length;
}

expressions.filters.findingCount = function(input, severity) {
    findings = count_findings_per_severity(input, severity);
    return (findings > 0) ? findings : '-';
}

expressions.filters.findingIdLabel = function(input, severity) {
    count = count_findings_per_severity(input, severity)
    if (count == 0)
        return '-';
    if (severity == 'None') {
        chap = 5;
        base = 1;
    } else {
        chap = 4;
        switch (severity) {
            case 'Critical':
                base = 1
                break;
            case 'High':
                base = count_findings_per_severity(input, 'Critical') + 1;
                break;
            case 'Medium':
                base = count_findings_per_severity(input, 'Critical') + 
                       count_findings_per_severity(input, 'High') + 1;
                break;
            case 'Low':
                base = count_findings_per_severity(input, 'Critical') + 
                       count_findings_per_severity(input, 'High') + 
                       count_findings_per_severity(input, 'Medium') + 1;
                break;
        }
    }
    switch (count) {
        case 1:
            return `${chap}.${base}`
        case 2:
            return `${chap}.${base} & ${chap}.${base+1}`
        default:
            return `${chap}.${base} t/m ${chap}.${base+count-1}`
    }
}

expressions.filters.findingCountSolved = function(input, severity) {
    if (count_findings_per_severity(input, severity) == 0)
        return '-';
    return input.filter((finding) => finding.cvss['environmentalSeverity'] == severity && finding.opgelost == 'Ja').length;
}

expressions.filters.spanTo = function(start, end, locale) {
    const start_date = new Date(start);
    const end_date = new Date(end);

    if (start_date == "Invalid Date" || end_date == "Invalid Date") return start;

    diff = Math.ceil(Math.abs(end_date - start_date) / (1000 * 60 * 60 * 24)) + 1; 
    str = (diff == 1) ? "{0} day" : "{0} days";
    return translate.translate(str).format((diff <= 20) ? numbers[locale][diff] : diff)
}

// Count vulnerability by category
// Example: {findings | countCategory: 'MyWebCategory'}
expressions.filters.countCategory = function(input, category) {
    if(!input) return input;
    var count = 0;

    for(var i = 0; i < input.length; i++){
        if(input[i].category === category){
            count += 1;
        }
    }
    return count;
}


// Convert input CVSS criteria into French: {input | criteriaFR}
expressions.filters.criteriaFR = function(input) {
    var pre = '<w:p><w:r><w:t>';
    var post = '</w:t></w:r></w:p>';
    var result = "Non défini"

    if (input === "Network") result = "Réseau"
    else if (input === "Adjacent Network") result = "Réseau Local"
    else if (input === "Local") result = "Local"
    else if (input === "Physical") result = "Physique"
    else if (input === "None") result = "Aucun"
    else if (input === "Low") result = "Faible"
    else if (input === "High") result = "Haute"
    else if (input === "Required") result = "Requis"
    else if (input === "Unchanged") result = "Inchangé"
    else if (input === "Changed") result = "Changé"

    // return pre + result + post;
    return result;
}

// Convert input date with parameter s (full,short): {input | convertDate: 's'}
expressions.filters.convertDateFR = function(input, s) {
    var date = new Date(input);
    if (date !== "Invalid Date") {
        var monthsFull = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
        var monthsShort = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];
        var days = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
        var day = date.getUTCDate();
        var month = date.getUTCMonth();
        var year = date.getUTCFullYear();
        if (s === "full") {
            return days[date.getUTCDay()] + " " + (day<10 ? '0'+day: day) + " " + monthsFull[month] + " " + year;
        }
        if (s === "short") {
            return (day<10 ? '0'+day: day) + "/" + monthsShort[month] + "/" + year;
        }
    }
}

// Convert input CVSS criteria into Spanish: {input | criteriaES}
expressions.filters.criteriaES = function(input) {
    var pre = '<w:p><w:r><w:t>';
    var post = '</w:t></w:r></w:p>';
    var result = "No definido"

    if (input === "Network") result = "Red"
    else if (input === "Adjacent Network") result = "Red adyacente"
    else if (input === "Local") result = "Local"
    else if (input === "Physical") result = "Físico"
    else if (input === "Required") result = "Obligatorio"
    else if (input === "Unchanged") result = "Sin cambiar"
    else if (input === "Changed") result = "Cambiado"
    else if (input === "Critical") result = "Crítica"
    else if (input === "Medium") result = "Media"
    else if (input === "None") result = "Informativa"
    else if (input === "Low") result = "Baja"
    else if (input === "High") result = "Alta"

    // return pre + result + post;
    return result;
}

// Convert input date with parameter s (full,short): {input | convertDateES: 's'}
expressions.filters.convertDateES = function(input, s) {
    var date = new Date(input);
    if (date !== "Invalid Date") {
        var monthsFull = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
        var monthsShort = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];
        var days = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"];
        var day = date.getUTCDate();
        var month = date.getUTCMonth();
        var year = date.getUTCFullYear();
        if (s === "full") {
            return days[date.getUTCDay()] + " " + (day<10 ? '0'+day: day) + " " + monthsFull[month] + " " + year;
        }
        if (s === "short") {
            return (day<10 ? '0'+day: day) + "/" + monthsShort[month] + "/" + year;
        }
         if (s === "OnlyYear") {
            return year;
        }
    }
}



exports.expressions = expressions

