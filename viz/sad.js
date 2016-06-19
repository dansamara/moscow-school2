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
	    map.geoObjects.add(
		new ymaps.Placemark(
		    [item.latitude, item.longitude],
		    {
			balloonContentBody: item.address,
		    },
		    {
			iconLayout: 'default#image',
		    	iconImageHref: 'i/program-0-1.png',
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


viz_map();
viz_images();
