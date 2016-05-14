
function viz_rating() {
    var container = d3.select('#rating-spark');
    if (container.size() == 0) {
	return;
    }

    var rating_2013 = +container.attr('data-rating-2013');
    var rating_2014 = +container.attr('data-rating-2014');
    var rating_2015 = +container.attr('data-rating-2015');

    if (rating_2013 && rating_2014 && rating_2015) {
	var data = [rating_2013, rating_2014, rating_2015];
	var years = [2013, 2014, 2015];
	var year = d3.select('#rating-year');
	var place = d3.select('#rating-place');

	var margin = {top: 3, right: 5, bottom: 4, left: 3};
	var width = 40 - margin.left - margin.right;
	var height = 15 - margin.top - margin.bottom;

	var svg = container.append('svg')
	    .attr('width', width + margin.left + margin.right)
	    .attr('height', height + margin.top + margin.bottom)
	    .append('g')
	    .attr('transform',
		  'translate(' + margin.left + ',' + margin.top + ')');

	var x = d3.scale.linear()
	    .range([0, width])
	    .domain([0, 2]);

	var y = d3.scale.linear()
	    .range([0, height])
	    .domain(d3.extent([rating_2013, rating_2014, rating_2015]));

	var line = d3.svg.line()
	    .x(function(d, i) { return x(i); })
	    .y(function(d) { return y(d); });

	svg.append('path')
	    .datum(data)
	    .attr('class', 'line')
	    .attr('d', line);

	var circles = svg.selectAll('circle')
	var bullets = circles
	    .data(data)
	    .enter().append('circle')
	    .attr('class', 'bullet')
	    .attr('cx', function(d, i) { return x(i); })
	    .attr('cy', function(d) { return y(d); });
	bullets = bullets[0];	// i do not know what i am doing

	function update(active) {
	    for (var index = 0; index < 3; index++) {
		var bullet = d3.select(bullets[index]);
		if (index == active) {
		    bullet.attr('r', 3)
		} else {
		    bullet.attr('r', 1)
		}
		year.text(years[active]);
		place.text(data[active]);
	    }
	}
	update(2);
	
	circles
	    .data(data)
	    .enter().append('circle')
	    .style('opacity', 0)
	    .attr('cx', function(d, i) { return x(i); })
	    .attr('cy', function(d) { return y(d); })
	    .attr('r', 12)
	    .on('mouseover', function(d, i) { update(i); })
	    .on('mouseout', function(d, i) { update(2); });
    }
}


function viz_map() {
    var data = [];
    var container = d3.select('#map ul');
    container.selectAll('li')
	.each(function() {
	    var item = d3.select(this);
	    data.push({
		address: item.text(),
		latitude: item.attr('data-latitude'),
		longitude: item.attr('data-longitude'),
		program: item.attr('data-program')
	    });
	});
    container.remove()

    function init() {
	var map = new ymaps.Map ('map', {
	    center: [55.76, 37.64],
	    zoom: 13,
	    controls: ['zoomControl']
	});
	map.behaviors.disable('scrollZoom');
	data.forEach(function(item) {
	    var program = item.program;
	    var icon;
	    if (program == 'program_1_4') {
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
	    }
	    map.geoObjects.add(
		new ymaps.Placemark(
		    [item.latitude, item.longitude],
		    {
			balloonContentBody: item.address + '. ' + program + '.',
		    },
		    {
			iconLayout: 'default#image',
		    	iconImageHref: icon,
		    	iconImageSize: [30, 30],
		    	iconImageOffset: [-15, -30] 
		    }
		)
	    );
	});
	if (data.length == 1) {
	    var item = data[0];
	    map.setCenter([item.latitude, item.longitude]);
	} else {
	    var bounds = map.geoObjects.getBounds();
	    map.setBounds(bounds);
	    map.setZoom(map.getZoom() - 1);
	}
    }
    ymaps.ready(init);
}


function viz_ege() {
    var container = d3.select('#ege-chart');
    if (container.size() == 0) {
	return
    }
    var a = +container.attr('data-a')
    var b = +container.attr('data-b')
    var legend = container.attr('data-legend')

    function distribution(a, b) {
	var data = [];
	var total = 0;
	for (var index = 0; index <= 300; index += 1) {
	    var p = index / 300;
	    var y = Math.pow(1 - p, a) * Math.pow(p, b)
	    total += y;
	    data.push(y);
	}
	return data.map(function(d) {
	    return d / total
	});
    }
    var data = distribution(a, b);
    var moscow = distribution(2.8, 5.2);

    var margin = {top: 5, right: 20, bottom: 45, left: 5};
    var width = 300 - margin.left - margin.right;
    var height = 180 - margin.top - margin.bottom;

    var table = container.append('table')
    var row = table.append('tr')
    var svg = row.append('td').append('svg')
	.attr('width', width + margin.left + margin.right)
	.attr('height', height + margin.top + margin.bottom)
	.append('g')
	.attr('transform',
	      'translate(' + margin.left + ',' + margin.top + ')');

    var x = d3.scale.linear()
	.domain([0, 300])
	.range([0, width]);

    var max = d3.max([
	d3.max(data),
	d3.max(moscow)
    ])
    var y = d3.scale.linear()
	.domain([0, max])
	.range([height, 0]);

    var xAxis = d3.svg.axis()
	.scale(x)
	.tickValues([0, 150, 220, 300])
	.orient('bottom');

    svg.append('g')
	.attr('class', 'axis')
	.attr('transform', 'translate(0,' + height + ')')
	.call(xAxis)
	.append('text')
	.attr('x', width - 20)
	.attr('y', 40)
	.style('text-anchor', 'end')
	.text('Сумма баллов за 3 экзамена');

    svg.append('g')
    	.append('text')
	.attr('transform', 'rotate(-90)')
	.attr('y', 10)
	.attr('x', -10)
	.style('text-anchor', 'end')
	.text('Доля учеников');

    var line = d3.svg.line()
	.x(function(d, i) { return x(i); })
	.y(function(d) { return y(d) - 2; });

    svg.append('path')
	.datum(data)
	.attr('class', 'area school')
	.attr('d', line);

    svg.append('path')
	.datum(moscow)
	.attr('class', 'area moscow')
	.attr('d', line);

    table = row.append('td').append('table')
    row = table.append('tr')
    row.append('td')
	.append('span')
	.attr('class', 'box school')
    row.append('td')
	.text(legend)
    row = table.append('tr')
    row.append('td')
	.append('span')
	.attr('class', 'box moscow')
    row.append('td')
	.text('В среднем по Москве')
}


function viz_olympiads() {
    var subjects = d3.selectAll('#olympiads > ol > li')
    var data = [];
    subjects[0]
	.forEach(function(d) {
	    var item = d3.select(d);
	    subject = item.select('span').text();
	    var types = []
	    var items = item.selectAll('li');
	    items[0]
		.forEach(function(d) {
		    var item = d3.select(d);
		    types.push({
			type: item.attr('data-type'),
			count: +item.text()
		    });
		});
	    data.push({
		subject: subject,
		types: types
	    });
	});
    var container = d3.select('#olympiads')
    container.select('ol').remove();
    var row = container.append('div')
	.attr('class', 'row')
    var tables = [
	row.append('div')
	    .attr('class', 'col-md-4')
	    .append('table'),
	row.append('div')
	    .attr('class', 'col-md-4')
	    .append('table'),
	row.append('div')
	    .attr('class', 'col-md-4')
	    .append('table')
    ]

    data.forEach(function(item, index) {
	var table = tables[index % 3];
	var row = table.append('tr');
	row.append('td')
	    .text(item.subject);
	var cell = row.append('td');
	item.types.forEach(function(item) {
	    var type = item.type;
	    var icon;
	    if (type == 'olympiad_moscow_1') {
		type = 'moscow-1';
	    } else if (type == 'olympiad_moscow_2') {
		type = 'moscow-2';
	    } else if (type == 'olympiad_russia_1') {
		type = 'russia-1';
		icon = 'i/gold-star.png';
	    } else if (type == 'olympiad_russia_2') {
		type = 'russia-2';
		icon = 'i/silver-star.png';
	    }
	    for (var index = 0; index < item.count; index += 1) {
		var box = cell.append('span')
		    .attr('class', 'box ' + type)
		if (icon) {
		    box.append('img')
			.attr('src', icon);
		}
	    }
	});
    });
}


function viz_universities() {
    var container = d3.select('#universities');
    if (container.size() == 0) {
	return
    }
    var list = container.select('ol');
    var type = list.attr('data-type');
    var universities = list.selectAll('li');
    data = []
    universities[0]
	.forEach(function(d) {
	    var item = d3.select(d);
	    var university = item.text();
	    if (type == 'vk') {
		var ids = item.attr('data-ids')
		if (ids) {
		    ids = ids.split(',');
		}
		var otmaza = item.attr('data-otmaza')
		data.push({
		    university: university,
		    ids: ids,
		    otmaza: otmaza
		});
	    } else if (type == 'manual') {
		var count = +item.attr('data-count');
		data.push({
		    university: university,
		    count: count
		});
	    }
	});

    list.remove();
    var row = container.append('div')
	.attr('class', 'row')
    var tables = [
	row.append('div')
	    .attr('class', 'col-md-4')
	    .append('table'),
	row.append('div')
	    .attr('class', 'col-md-4')
	    .append('table'),
	row.append('div')
	    .attr('class', 'col-md-4')
	    .append('table')
    ]
    if (type == 'vk') {
	data.forEach(function(item, index) {
	    var table = tables[index % 3];
	    var row = table.append('tr');
	    var university = item.university
	    row.append('td')
		.text(university);
	    var cell = row.append('td');
	    if (item.ids) {
		item.ids.forEach(function(id) {
		    cell.append('a')
			.attr('href', 'https://vk.com/id' + id)
			.attr('class', 'box student')
		});
	    } else {
		cell.append('abbr')
		    .attr('title', item.otmaza)
		    .text('?')
	    }
	});
    } else if (type == 'manual') {
	data.forEach(function(item, index) {
	    var table = tables[index % 3];
	    var row = table.append('tr');
	    var university = item.university
	    row.append('td')
		.text(university);
	    var cell = row.append('td');
	    var count = item.count;
	    for (var index = 0; index < count; index += 1) {
		cell.append('span')
		    .attr('class', 'box student')
	    }
	});
    }
}


function viz_teachers() {
    var subjects = d3.selectAll('#teachers > ol > li')
    var data = [];
    subjects[0]
	.forEach(function(d) {
	    var item = d3.select(d);
	    subject = item.select('span').text();
	    var categories = []
	    var items = item.selectAll('li a');
	    items[0]
		.forEach(function(d) {
		    var item = d3.select(d);
		    url = item.attr('href');
		    if (url == 'email') {
			url = null;
		    }
		    categories.push({
			url: url,
			category: item.text()
		    });
		});
	    data.push({
		subject: subject,
		categories: categories
	    });
	});

    var container = d3.select('#teachers')
    container.select('ol').remove();
    var row = container.append('div')
	.attr('class', 'row')
    var tables = [
	row.append('div')
	    .attr('class', 'col-md-6')
	    .append('table'),
	row.append('div')
	    .attr('class', 'col-md-6')
	    .append('table'),
    ]

    data.forEach(function(item, index) {
	var table = tables[index % 2];
	var row = table.append('tr');
	row.append('td')
	    .text(item.subject);
	var cell = row.append('td');
	item.categories.forEach(function(item) {
	    var category = item.category;
	    var url = item.url;
	    if (url) {
		var box = cell.append('a')
		    .attr('class', 'box category-' + category)
		    .attr('href', url);
	    } else {
		var box = cell.append('span')
		    .attr('class', 'box category-' + category)
	    }
	    if (category == 'high') {
		box.append('img')
		    .attr('src', 'i/high-star.png')
	    } else if (category == '1') {
		box.append('span')
		    .text('1')
	    }
	});
    });    
}


function viz_images() {
    var container = d3.select('#images');
    if (container.size() == 0) {
	return
    }
    var links = container.selectAll('a');
    var data = [];
    links[0]
	.forEach(function(d) {
	    var link = d3.select(d);
	    var big = link.attr('href');
	    var size = link.attr('data-size')
		.split('x');
	    var width = +size[0];
	    var height = +size[1];
	    var small = link.select('img').attr('src');
	    data.push({
		big: big,
		small: small,
		aspect_ratio: width / height
	    });
	});
    links.remove();
    $("#images").chromatic(data);
}


function viz_teacher_pupils() {
    var container = d3.select('#pupils-teachers');
    if (container.size() == 0) {
	return
    }
    var share = +container.attr('data-share');
    var distribution = {4: 1,
			7: 7,
			8: 10,
			9: 34,
			10: 122,
			11: 155,
			12: 102,
			13: 57,
			14: 22,
			15: 13,
			16: 6,
			17: 5,
			18: 1,
			19: 1,
			20: 2,
			22: 1,
			23: 2,
			27: 1};
    var data = [];
    for (var key in distribution) {
	data.push({
	    key: +key,
	    value: distribution[key]
	});
    }

    var margin = {top: 5, right: 20, bottom: 25, left: 5};
    var width = 300 - margin.left - margin.right;
    var height = 100 - margin.top - margin.bottom;

    var svg = container.append('svg')
    	.attr('width', width + margin.left + margin.right)
    	.attr('height', height + margin.top + margin.bottom)
    	.append('g')
    	.attr('transform',
    	      'translate(' + margin.left + ',' + margin.top + ')');

    var extent = d3.extent(
	data,
	function(d) { return d.key; }
    );
    var x = d3.scale.linear()
    	.domain(extent)
    	.range([0, width]);

    var max = d3.max(
	data,
	function(d) { return d.value; }
    )
    var y = d3.scale.linear()
    	.domain([0, max])
    	.range([height, 0]);

    var xAxis = d3.svg.axis()
    	.scale(
	    d3.scale.linear()
		.domain([9, 15])
		.range([x(9), x(15)])
	)
	.tickFormat(d3.format('d'))
	.tickValues([9, 11, 15])
    	.orient('bottom');

    svg.append('g')
    	.attr('class', 'axis')
    	.attr('transform', 'translate(0,' + height + ')')
    	.call(xAxis);

    var step = x(2) - x(1);
    svg.selectAll('.bar')
	.data(data)
	.enter().insert('rect', '.axis')
	.attr('class', 'bar')
	.attr('x', function(d) { return x(d.key); })
	.attr('y', function(d) { return y(d.value * 0.75) - 2; })
	.attr('width', step)
	.attr('height', function(d) { return height - y(d.value * 0.75); });

    svg.append('line')
	.attr('class', 'share')
	.attr({
	    x1: x(share),
	    y1: y(height * 1.7),
	    x2: x(share),
	    y2: y(0) - 2
	});

    svg.append('text')
	.attr('x', x(share))
    	.attr('y', y(height * 1.7) - 5)
    	.style('text-anchor', 'middle')
    	.text(share);

}


function viz_salaries() {
    var container = d3.select('#salaries');
    if (container.size() == 0) {
	return
    }
    var salary = +container.attr('data-salary');
    var distribution = {40000: 1,
			55000: 1,
			56000: 1,
			57000: 1,
			58000: 2,
			62000: 1,
			63000: 1,
			64000: 6,
			65000: 6,
			66000: 10,
			67000: 9,
			68000: 20,
			69000: 20,
			70000: 37,
			71000: 35,
			72000: 44,
			73000: 54,
			74000: 49,
			75000: 48,
			76000: 42,
			77000: 31,
			78000: 30,
			79000: 17,
			80000: 17,
			81000: 10,
			82000: 12,
			83000: 4,
			84000: 6,
			85000: 5,
			86000: 6,
			87000: 4,
			88000: 3,
			89000: 1,
			90000: 4,
			91000: 1,
			94000: 1,
			95000: 1,
			106000: 1};
    var data = [];
    for (var key in distribution) {
	data.push({
	    key: +key,
	    value: distribution[key]
	});
    }

    var margin = {top: 5, right: 40, bottom: 25, left: 40};
    var width = 330 - margin.left - margin.right;
    var height = 100 - margin.top - margin.bottom;

    var svg = container.append('svg')
    	.attr('width', width + margin.left + margin.right)
    	.attr('height', height + margin.top + margin.bottom)
    	.append('g')
    	.attr('transform',
    	      'translate(' + margin.left + ',' + margin.top + ')');

    var extent = d3.extent(
	data,
	function(d) { return d.key; }
    );
    var x = d3.scale.linear()
    	.domain(extent)
    	.range([0, width]);

    var max = d3.max(
	data,
	function(d) { return d.value; }
    )
    var y = d3.scale.linear()
    	.domain([0, max])
    	.range([height, 0]);

    function format_salary(salary) {
	var string = salary.toString();
	return (string.substring(0, string.length - 3)
		+ ' ' + string.substring(string.length - 3));
    }

    var xAxis = d3.svg.axis()
    	.scale(
	    d3.scale.linear()
		.domain([60000, 90000])
		.range([x(60000), x(90000)])
	)
    	.tickValues([60000, 75000, 90000])
	.tickFormat(format_salary)
    	.orient('bottom');

    svg.append('g')
    	.attr('class', 'axis')
    	.attr('transform', 'translate(0,' + height + ')')
    	.call(xAxis);

    var step = x(1000) - x(0);
    svg.selectAll('.bar')
	.data(data)
	.enter().insert('rect', '.axis')
	.attr('class', 'bar')
	.attr('x', function(d) { return x(d.key); })
	.attr('y', function(d) { return y(d.value * 0.75) - 2; })
	.attr('width', step)
	.attr('height', function(d) { return height - y(d.value * 0.75); });

    svg.append('line')
	.attr('class', 'salary')
	.attr({
	    x1: x(salary),
	    y1: y(height / 1.7),
	    x2: x(salary),
	    y2: y(0) - 2
	});

    svg.append('text')
	.attr('x', x(salary))
    	.attr('y', y(height / 1.7) - 5)
    	.style('text-anchor', 'middle')
    	.text(format_salary(salary) + ' руб.');
}


viz_rating();
viz_map();
viz_ege();
viz_olympiads();
viz_universities();
viz_teachers();
viz_images();
viz_teacher_pupils();
viz_salaries();
