function getState() {
    var path = window.location.search;
    var thanks = /thanks/.exec(path) !== null;
    var match = /school=([^&$]+)/.exec(path);
    var school;
    if (match !== null) {
	school = match[1];
    }
    return {
	'thanks': thanks,
	'school': school
    }
}


var state = getState();
if (state.thanks) {
    $('#thanks').css('display', 'block');
}
var school = state.school;
if (school) {
    $('select[name="review-school"]').val(school);
    $('select[name="bug-school"]').val(school);
}
