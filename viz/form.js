function getState() {
    var path = window.location.search;
    var thanks = /thanks/.exec(path) !== null;
    var match = /(school|sad)=([^&$]+)/.exec(path);
    var group;
    var school_sad;
    if (match !== null) {
	group = match[1];
	school_sad = match[2];
    }
    return {
	'thanks': thanks,
	'group': group,
	'school_sad': school_sad
    }
}


var state = getState();
if (state.thanks) {
    $('#thanks').css('display', 'block');
}
var school_sad = state.school_sad;
if (school_sad) {
    var group = state.group;
    if (group == 'school') {
	$('select[name="review-school"]').val(school_sad);
    } else if (group == 'sad') {
	$('select[name="review-sad"]').val(school_sad);
    }
    $('select[name="bug-school-sad"]').val(school_sad);
}
