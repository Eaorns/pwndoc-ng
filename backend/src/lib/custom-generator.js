var expressions = require('angular-expressions');

// Apply all customs functions
function apply(data) {

}
exports.apply = apply;

// *** Custom modifications of audit data for usage in word template


// *** Custome Angular expressions filters ***

var filters = {};

expressions.filters.cveSlider = function(input, score) {
    if (score == undefined || score < 0 || score > 10)
        score = 0;

    return "/app/images/sliders/Slider_groen-rood_" + Math.round(score)*10 + ".png";
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

