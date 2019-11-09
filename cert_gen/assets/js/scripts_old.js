
function scroll_to_class(chosen_class) {
	var nav_height = $('nav').outerHeight();
	var scroll_to = $(chosen_class).offset().top - nav_height;

	if($(window).scrollTop() != scroll_to) {
		$('html, body').stop().animate({scrollTop: scroll_to}, 1000);
	}
}


jQuery(document).ready(function() {

	/*
	    Fullscreen background
	*/
	$.backstretch("assets/img/backgrounds/1.jpg");

	/*
	    Multi Step Form
	*/
	$('.msf-form form fieldset:first-child').fadeIn('slow');
	
	// next step
	$('.msf-form form .btn-next').on('click', function() {
		$(this).parents('fieldset').fadeOut(400, function() {
	    	$(this).next().fadeIn();
			scroll_to_class('.msf-form');
	    });
	});
	
	// previous step
	$('.msf-form form .btn-previous').on('click', function() {
		$(this).parents('fieldset').fadeOut(400, function() {
			$(this).prev().fadeIn();
			scroll_to_class('.msf-form');
		});
	});
	
	//submit button
	//$('.msf-form form .submit').on('submit', function() {
	
});

// Form Validation




//////////////////////////////
/***** Python Functions *****/
//////////////////////////////

function getFormData($form){

    console.log("Creating JSON");
    var obj = $('form').serializeJSON();
    var jsonString = JSON.stringify(obj);

    console.log("Passing to Python");
    //python
    eel.main_func(jsonString)();
    window.close()
}

// Ask the user for a file using Python
async function getFile(for_id, type) {
    document.getElementById(for_id).value = await eel.ask_file(type)();
}

