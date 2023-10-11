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
    score = typeof score == 'object' ? score.value : score;
    score = typeof score == "String" ? parseInt(score) : score;
    console.log("SCORE", score);

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
    // source: https://gist.github.com/aminnj/5ca372aa2def72fb017b531c894afdca
    char_weights = {
        " ": 4.4453125,  "!": 4.4453125,  '"': 5.6796875, 
        "#": 8.8984375,  "$": 8.8984375,  "%": 14.2265625, 
        "&": 10.671875,  "'": 3.0546875,  "(": 5.328125, 
        ")": 5.328125,   "*": 6.2265625,  "+": 9.34375, 
        ",": 4.4453125,  "-": 5.328125,   ".": 4.4453125, 
        "/": 4.4453125,  "0": 8.8984375,  "1": 7.7228125, 
        "2": 8.8984375,  "3": 8.8984375,  "4": 8.8984375, 
        "5": 8.8984375,  "6": 8.8984375,  "7": 8.8984375, 
        "8": 8.8984375,  "9": 8.8984375,  ":": 4.4453125, 
        ";": 4.4453125,  "<": 9.34375,    "=": 9.34375, 
        ">": 9.34375,    "?": 8.8984375,  "@": 16.2421875, 
        "A": 10.671875,  "B": 10.671875,  "C": 11.5546875, 
        "D": 11.5546875, "E": 10.671875,  "F": 9.7734375, 
        "G": 12.4453125, "H": 11.5546875, "I": 4.4453125, 
        "J": 8,          "K": 10.671875,  "L": 8.8984375, 
        "M": 13.328125,  "N": 11.5546875, "O": 12.4453125, 
        "P": 10.671875,  "Q": 12.4453125, "R": 11.5546875, 
        "S": 10.671875,  "T": 9.7734375,  "U": 11.5546875, 
        "V": 10.671875,  "W": 15.1015625, "X": 10.671875, 
        "Y": 10.671875,  "Z": 9.7734375,  "[": 4.4453125, 
       "\\": 4.4453125,  "]": 4.4453125,  "^": 7.5078125, 
        "_": 8.8984375,  "`": 5.328125,   "a": 8.8984375, 
        "b": 8.8984375,  "c": 8,          "d": 8.8984375, 
        "e": 8.8984375,  "f": 4.15921875, "g": 8.8984375, 
        "h": 8.8984375,  "i": 3.5546875,  "j": 3.5546875, 
        "k": 8,          "l": 3.5546875,  "m": 13.328125, 
        "n": 8.8984375,  "o": 8.8984375,  "p": 8.8984375, 
        "q": 8.8984375,  "r": 5.328125,   "s": 8, 
        "t": 4.4453125,  "u": 8.8984375,  "v": 8, 
        "w": 11.5546875, "x": 8,          "y": 8, 
        "z": 8,          "{": 5.34375,    "|": 4.15625, 
        "}": 5.34375,    "~": 9.34375
    }

    let calc_width = (s) => s.split('').map((c) => char_weights[c]).reduce((p, v) => p + v, 0)

    num_vals = input.length;
    vals = input.map((x) => [x, calc_width(x)]).sort((a, b) => a[1] - b[1]);
    console.log(vals);

    // bins = [[], [], [], [], [], []];
    // overflow = [];
    // for (val of vals) {
    //     // First, approximate
    //     bin = Math.ceil(val[1] / (col_char_limit + bet_char_limit));
    //     // Then, check if actually correct or should be adjusted
    //     if (bin * col_char_limit + (bin-1) * bet_char_limit < val[1])
    //         bin++;
    //     if (bin >= 7)
    //         overflow.push(val[0]);
    //     else
    //         bins[bin-1].push(val[0]);
    // }
    // // bins = bins.reverse();
    // console.log(bins);
    // console.log(overflow);

    // num_in_bins = num_vals - overflow.length;
    // while (num_in_bins > 0) {
    //     cols_left = 6;
    //     cols_tried = cols_left;
    //     row = [];
    //     while (cols_left > 0 && cols_tried > 0) {
    //         if (bins[cols_tried-1].length > 0) {
    //             row.push([bins[cols_tried-1].pop(), cols_tried]);
    //             cols_left -= cols_tried;
    //             cols_tried = cols_left;
    //             num_in_bins--;
    //         } else {
    //             cols_tried--;
    //         }
    //     }
    //     if (cols_left > 0) {
    //         row[row.length-1][1] += cols_left;
    //     }
    // }


    median = vals[Math.floor(vals.length*.6)][1]; // Not 'real' median, slight offset
    console.log("median", median);
    var num_cols = 6;
    bet_char_limit = char_weights["'"]*6
    max_char_limit = char_weights["'"]*238
    col_char_limit = 0;
    for (; num_cols > 1; num_cols--) {
        col_char_limit = (max_char_limit - bet_char_limit * (num_cols-1)) / num_cols;
        if (col_char_limit > median)
            break;
    }

    console.log("COLS", num_cols, col_char_limit, bet_char_limit);
    vals = vals.map(function(val) {
        // First, approximate
        colspan = Math.ceil(val[1] / (col_char_limit + bet_char_limit));
        // Then, check if actually correct or should be adjusted
        if (colspan * col_char_limit + (colspan-1) * bet_char_limit < val[1])
            colspan++;
        return [val[0], Math.min(colspan, num_cols)];
    });
 
    console.log(vals);
    rows = [];
    while (vals.length > 0) {
        cols_left = num_cols;
        row = []
        while(cols_left > 0 && vals.length > 0) {
            if (vals[0][1] <= cols_left) {
                cols_left -= vals[0][1]
                row.push(vals.shift())
            } else {
                row[row.length-1][1] += cols_left;
                break;
            }
        }
        if (vals.length == 0)
            row[row.length-1][1] += cols_left;

        rows.push(row);
    }
    console.log(rows);

    tw = 5100;
    cw = tw / num_cols;

    // https://docxperiments.readthedocs.io/en/latest/synthesis/documentxml.html
    pre = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid" /><w:tblW w:w="${tw}" w:type="pct" /><w:tblBorders><w:top w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:left w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:bottom w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:right w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideH w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideV w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /></w:tblBorders><w:tblLayout w:type="fixed" /></w:tblPr>`;
    pre += '<w:tblGrid>' + ` <w:gridCol w:w="${cw}" />`.repeat(num_cols) + '</w:tblGrid>';
    pre += `<w:tr><w:tc><w:tcPr><w:tcW w:w="${tw}" w:type="dxa" /><w:gridSpan w:val="${num_cols}" /><w:shd w:val="clear" w:color="auto" w:fill="2DA5DE" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF" w:themeColor="background1" /></w:rPr><w:t xml:space="preserve">Doelobjecten</w:t></w:r></w:p></w:tc></w:tr>`
    post = '</w:tbl>';

    out = ""
    for (row of rows) {
        out += "<w:tr>";
        for (col of row) {
            console.log(cw*col[1], col[1], col[0]);
            out += `<w:tc><w:tcPr><w:tcW w:w="${cw*col[1]}" w:type="dxa" /><w:gridSpan w:val="${col[1]}" /><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1" w:themeFillShade="D9" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:t>${col[0]}</w:t></w:r></w:p></w:tc>`;
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

function generate_table(input, col_width, col_names) {
    tw = col_width.reduce((v, a) => v + a, 0);

    // https://docxperiments.readthedocs.io/en/latest/synthesis/documentxml.html
    pre = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid" /><w:tblW w:w="${tw}" w:type="pct" /><w:tblBorders><w:top w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:left w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:bottom w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:right w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideH w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /><w:insideV w:val="single" w:sz="24" w:space="0" w:color="FFFFFF" w:themeColor="background1" /></w:tblBorders><w:tblLayout w:type="fixed" /></w:tblPr>`;
    pre += '<w:tblGrid>' + col_width.map((w) => ` <w:gridCol w:w="${w}" />`).join('') + '</w:tblGrid>';
    pre += '<w:tr>' + col_width.map((w, idx) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa" /><w:shd w:val="clear" w:color="auto" w:fill="2DA5DE" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF" w:themeColor="background1" /><w:b w:val="true"/></w:rPr><w:t xml:space="preserve">${col_names[idx]}</w:t></w:r></w:p></w:tc>`).join('') + '</w:tr>' 
    post = '</w:tbl>';
    
    out = "";
    for (row of input) {
        out += "<w:tr>";
        for (var i = 0; i < input[0].length; i++) {
            data = html2ooxml(row[i].replace(/(<p><\/p>)+$/, '')).replaceAll(/(^<w:p>|<\/w:p>$)/gm, '')
            out += `<w:tc><w:tcPr><w:tcW w:w="${col_width[i]}" w:type="dxa" /><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9" w:themeFill="background1" w:themeFillShade="D9" /></w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto" /></w:pPr>${data}</w:p></w:tc>`
        }
        out += "</w:tr>";
    }

    return pre + out + post;
}

expressions.filters.generateOutdatedSoftwareTable = function(input) {
    if (input == null || input.length == 0)
        return '';

    switch (input[0].length) {
        case 2:
            var col_width = [1200, 2800];
            var col_names = ['Doelobjecten', 'Geïnstalleerde versie/Oplossing'];
            break;
        case 3:
            var col_width = [1200, 2100, 1800];
            var col_names = ['Kwetsbaar product', 'Geïnstalleerde versie/Oplossing', 'Doelobjecten'];
            break;
        case 4:
            var col_width = [1200, 1800, 600, 1500];
            var col_names = ['Kwetsbaar product', 'Geïnstalleerde versie/Oplossing', 'CVSS', 'Doelobjecten'];
            break;
        default:
            console.log(`Outdated Software table has invalid number of columns (${input[0].length})! not parsing. Data: ${input}`);
            return "";
    }
    
    return generate_table(input, col_width, col_names)
}

expressions.filters.generateGlobalTargetsTable = function(input) {
    return generate_table(input, [800, 1300, 500, 2500], ['IP-adres', 'systeemnaam / URL', 'IPv6', 'Omschrijving'])
}

expressions.filters.generateExceptionsTable = function(input) {
    return generate_table(input, [1800, 3300], ['Doelobject', 'Niet onderzocht in verband met'])
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
    return translate.translate(str, locale).format((diff <= 20) ? numbers[locale][diff] : diff)
}

expressions.filters.addChar = function(input, char) {
    return (input.slice(-1) == char) ? input : input + char;
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

