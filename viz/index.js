function init() {
    var data = [];
    $('ul#school-sads > li').each(function(index, school) {
	school = $(school);
	var link = school.find('a');
	var title = link.text();
	var url = link.attr('href');
	school.find('ul li').each(function(index, address) {
	    address = $(address);
	    var description = address.text();
	    var latitude = +address.attr('data-latitude');
	    var longitude = +address.attr('data-longitude');
	    var program = address.attr('data-program');
	    var building = +address.attr('data-building')
	    data.push({
		title: title,
		url: url,
		address: description,
		latitude: latitude,
		longitude: longitude,
		program: program,
		building: building
	    });
	});
    });
    $('ul#school-sads').remove()

    var map = new ymaps.Map ('map', {
	center: [55.76, 37.64],
	zoom: 12,
	controls: ['smallMapDefaultSet']
    });
    map.copyrights.add('<a href="http://lab.alexkuk.ru">Лаборатория анализа данных Кукушкина Александра</a>');
    map.copyrights.add('<a href="http://russianschools.ru">Российские школы</a>')

    data.forEach(function(item) {
	var icon;
	var program = item.program;
	if (program == 'program_0_1') {
	    icon = 'i/program-0-1.png';
	    program = null;
	} else if (program == 'program_1_4') {
	    icon = 'i/program-1-4.png';
	    program = 'Обучение с 1 по 4 класс';
	} else if (program == 'program_5_11') {
	    icon = 'i/program-5-11.png'
	    program = 'Обучение с 5 по 11 класс';
	} else if (program == 'program_6_11') {
	    icon = 'i/program-5-11.png'
	    program = 'Обучение с 6 по 11 класс';
	} else if (program == 'program_1_11') {
	    icon = 'i/program-1-11.png'
	    program = 'Обучение с 1 по 11 класс';
	} else if (program == 'program_1_9') {
	    icon = 'i/program-1-4.png'
	    program = 'Обучение с 1 по 9 класс';
	} else if (program == 'program_10_11') {
	    icon = 'i/program-5-11.png'
	    program = 'Обучение с 10 по 11 класс';
	} else if (program == 'program_8_11') {
	    icon = 'i/program-5-11.png'
	    program = 'Обучение с 8 по 11 класс';
	} else if (program == 'program_1_7') {
	    icon = 'i/program-1-4.png'
	    program = 'Обучение с 1 по 7 класс';
	}

	var balloon = (
	    '<div class="balloon">' +
		'<h1>' + item.title + '</h1>' +
		'<div>'
	);
	if (item.building > 0) {
	    balloon += '<span>Корпус № ' + item.building + '.</span> '
	}
	if (program) {
	    balloon += '<span>' + program + '.</span> '
	}
	balloon += (
	    '<span>' + item.address + '.</span></div>' +
		'<div class="more"><a href="' + item.url + '">' +
		'Посмотреть подробную информацию</a></div>' +
		'</div>'
	)

	map.geoObjects.add(
	    new ymaps.Placemark(
		[item.latitude, item.longitude],
		{
		    balloonContentBody: balloon,
		},
		{
		    iconLayout: 'default#image',
		    iconImageHref: icon,
		    iconImageSize: [20, 20],
		    iconImageOffset: [-10, -20] 
		}
	    )
	);
    });
}

ymaps.ready(init);

