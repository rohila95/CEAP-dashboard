/**
 * 
 */

function showRGDPGraph(region)
{
	$("#ceapChart").empty();
	$("#ceapChart").append("<h3 align=\"center\">Real Gross Domestic Product by Year</h3><h5 align=\"center\">(In Billions of Dollars)</h5>");
	console.log("Region is "+region);
		
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	// Set the ranges
	var	x = d3.scale.ordinal().rangeRoundBands([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom").ticks(5);
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).orient("left").ticks(5).tickFormat(function(d){
			return "$"+d3.format(",")(d);
		});
	
	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = "<strong>Value:</strong> <span style='color:white'>"
								+d.value + "</span>";
					});
	    
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
		.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	 
	svg.call(tip);
	
	// Get the data
	d3.csv("RGDP.csv", function(error, data) {
		data.forEach(function(d) {
				d.Year = d.Year;
				d.value = +d[region]/1000;
		});
	 
		// Scale the range of the data
		x.domain(data.map(function(d) { 
			if(parseInt(d.Year)>=2001)
			{
			return d.Year; }}));
		y.domain([0, d3.max(data, function(d) {
			if(parseInt(d.Year)>=2001)
			{
			return d.value;} })]);
	 
		 svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");

	  svg.selectAll("bar")
	      .data(data)
	    .enter().append("rect")
	      .style("fill", "#4dc3ff")
	      .attr("x", function(d) { 
				if(parseInt(d.Year)>=2001)
				{
	    	  return x(d.Year)+10; }})
	      .attr("width",35)
	      .attr("y", function(d) { 
				if(parseInt(d.Year)>=2001)
				{
	    	  return y(d.value); }})
	      .attr("height", function(d) {
				if(parseInt(d.Year)>=2001)
				{
	    	  return height - y(d.value)}})
	       .on('mouseover', tip.show).on('mouseout', tip.hide);
	 
	});
	
	
		if(region=="Virginia")
			{
			showVAGdp(region);
			showGrowthRateGDP_VA(region);
			showQuartGrowthVa(region);
			}
		else
			showGrowthRateGDP(region);
}

function showGrowthRateGDP(region)
{
	$("#ceapChart2").empty();
	$("#ceapChart2").append("<h3 align=\"center\">Growth Rate of Real GDP</h3>");
	console.log("Region is "+region);
		var prev=0;
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	// Set the ranges
	var	x = d3.scale.ordinal().rangeRoundBands([0, width]);
	var	y = d3.scale.linear().range([height, 0]).nice();
	var flag=1;
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom").ticks(5);
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).orient("left").ticks(5).tickFormat(function(d){
			return d3.format(",")(d)+"%";
		});
	
	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = "<strong>Value:</strong> <span style='color:white'>"
								+d.value + "</span>";
					});
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart2")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
		.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	 
	svg.call(tip);
	var newdata=[];
	// Get the data
	d3.csv("RGDP.csv", function(error, data) {
		data.forEach(function(d) {
			
			d.Year = d.Year;
			//d.value = +d[region]/1000;
			if(flag==1)
				{
				d.value=0;
				//console.log(prev+"-"+flag);
				flag=0;
				}
			else{
				d.value=(d[region]/(prev*1000))*100-100;
				//console.log(prev+"-"+flag);
				//console.log(d[region]+"-"+(prev*1000)+"----"+d.value+"-"+prev);
				console.log(d.value+"---"+d.Year);
				newdata.push(d)
			}
				prev=d[region]/1000;
				
				
		});
		// Scale the range of the data
		x.domain(newdata.map(function(d) { 
			if(parseInt(d.Year)>=2001)
			{
			return d.Year; }}));

		  y.domain(d3.extent(newdata.map(function(d) { return d.value;} )));
	
	  svg.selectAll("bar")
	      .data(newdata)
	    .enter().append("rect")
	      .style("fill", function(d, i) { return d.value < 0 ? "#e25846" : "#4dc3ff"; })
	      .attr("x", function(d,i) { 
				if(parseInt(d.Year)>=2001)
				{
					
	    	  return x(d.Year)+10; }})
	      .attr("width",35)
	      .attr("y", function(d) { 
				if(parseInt(d.Year)>=2001)
				{
	    	  return y(Math.max(0, d.value)); }})
	      .attr("height", function(d) {
				if(parseInt(d.Year)>=2001)
				{
	    	  return Math.abs(y(d.value) - y(0));}})
	       .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," + y(0) + ")")
      .call(xAxis)
    .selectAll("text")
      .style("text-anchor", "end")
      .attr("dx", "-.8em")
      .attr("dy", "-.55em")
      .attr("transform", "rotate(-90)" );

	 svg.append("g")
      .attr("class", "y axis")
      .call(yAxis)
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", ".71em")
      .style("text-anchor", "end")
      .text(" ");
	 
	});
}

function showGrowthRateGDP_VA(region)
{

	$("#ceapChart3").empty();
	$("#ceapChart3").append("<h3 align=\"center\">Growth Rate of Real GDP</h3>");
	console.log("Region is "+region);
		var prev=0;
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	// Set the ranges
	var	x = d3.scale.ordinal().rangeRoundBands([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	var flag=1;
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom").ticks(5);
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).orient("left").ticks(5).tickFormat(function(d){
			return d3.format(",")(d)+"%";
		});
	
	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = "<strong>Value:</strong> <span style='color:white'>"
								+d.value + "</span>";
					});
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart3")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
		.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	var newdata=[]; 
	svg.call(tip);
	
	// Get the data
	d3.csv("RGDP.csv", function(error, data) {
		data.forEach(function(d) {
			
			d.Year = d.Year;
			//d.value = +d[region]/1000;
			if(flag==1)
				{
				d.value=0;
				//console.log(prev+"-"+flag);
				flag=0;
				}
			else{
				d.value=(d[region]/(prev*1000))*100-100;
				//console.log(prev+"-"+flag);
			//	console.log(d[region]+"-"+(prev*1000)+"----"+d.value+"-"+prev);
				newdata.push(d);
				
			}
				prev=d[region]/1000;
			
				
		});
		// Scale the range of the data
		x.domain(newdata.map(function(d) { 
			if(parseInt(d.Year)>=2001)
			{
			return d.Year; }}));
//		y.domain([0, d3.max(data, function(d) {
//			if(parseInt(d.Year)>=2001)
//			{
//			return d.value;} })]);
		
		  y.domain(d3.extent(newdata.map(function(d) { return d.value;} )));
	 
		

	  svg.selectAll("bar")
	      .data(newdata)
	    .enter().append("rect")
	      .style("fill", function(d, i) { return d.value < 0 ? "#e25846" : "#4dc3ff"; })
	      .attr("x", function(d) { 
				if(parseInt(d.Year)>=2001)
				{
	    	  return x(d.Year)+10; }})
	      .attr("width",35)
	      .attr("y", function(d) { 
				if(parseInt(d.Year)>=2001)
				{
	    	  return y(Math.max(0, d.value));  }})
	      .attr("height", function(d) {
				if(parseInt(d.Year)>=2001)
				{
	    	  return Math.abs(y(d.value) - y(0));}})
	       .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," +  y(0) + ")")
      .call(xAxis)
    .selectAll("text")
      .style("text-anchor", "end")
      .attr("dx", "-.8em")
      .attr("dy", "-.55em")
      .attr("transform", "rotate(-90)" );

	 svg.append("g")
      .attr("class", "y axis")
      .call(yAxis)
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", ".71em")
      .style("text-anchor", "end")
      .text(" ");
	 
	});
	
}
function showQuartGrowthVa(region)
{
	$("#ceapChart4").empty();
	$("#ceapChart4").append("<h3 align=\"center\">Quarterly growth rate</h3>");
	console.log("Region is "+region);
	var regionNames={"Virginia":"avagdp"};
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
	width = 900 - margin.left - margin.right,
	height = 500 - margin.top - margin.bottom;
	
	var formatDate = d3.time.format("%m/%d/%Y").parse;
	
	// Set the ranges
	var	x = d3.time.scale().range([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom");
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).orient("left").ticks(5).orient("left").ticks(5).tickFormat(function(d){
			return d+"%";
		});
	
	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = "<span style='color:white'>"
							+d.tiplabel+" : "+d.value + "</span>";
					});
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart4")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
			.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	var newdata=[];
	svg.call(tip);
	//Get the data
	d3.csv("VA_gdp.csv", function(error, data) {
		data.forEach(function(d) {	
			if(d.date==null)
				d.date=0;
			d.label = formatDate(d.date);
			d.tiplabel = d.date;
			d.value = +d[regionNames[region]];
			//console.log(d.date+"--"+d.value);
			if(d.value!=0)
				newdata.push(d);
			
		});
	 
		// Scale the range of the data
		x.domain(d3.extent(newdata,function(d) {
			return d.label; 
			}));
		  y.domain(d3.extent(newdata.map(function(d) { return d.value;} )));

	  svg.selectAll("bar")
	      .data(newdata)
	      .enter().append("rect")
	      .style("fill", function(d, i) { return d.value < 0 ? "#e25846" : "#4dc3ff"; })
	      .attr("x", function(d) { 
	    	
	    	  return x(d.label); })
	      .attr("width",10)
	      .attr("y", function(d) {
	    	  return y(Math.max(0, d.value)); })
	      .attr("height", function(d) {
	    	 
	    	  return Math.abs(y(d.value) - y(0));})
	    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
	  .attr("class", "x axis")
	  .attr("transform", "translate(0," + y(0)  + ")")
	  .call(xAxis)
	.selectAll("text")
	  .style("text-anchor", "end")
	  .attr("dx", "-.8em")
	  .attr("dy", "-.55em")
	  .attr("transform", "rotate(-90)" );

	 svg.append("g")
	  .attr("class", "y axis")
	  .call(yAxis)
	.append("text")
	  .attr("transform", "rotate(-90)")
	  .attr("y", 6)
	  .attr("dy", ".71em")
	  .style("text-anchor", "end")
	  .text(" ");
	 
	});
}
function showVAGdp(region)
{
	$("#ceapChart2").empty();
	$("#ceapChart2").append("<h3 align=\"center\">Real Gross Domestic Product by Quarter</h3><h5 align=\"center\">(In Billions of Dollars)</h5>");
	console.log("Region is "+region);
	var regionNames={"Virginia":"rgdpva"};
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
	width = 900 - margin.left - margin.right,
	height = 500 - margin.top - margin.bottom;
	
	var formatDate = d3.time.format("%m/%d/%Y").parse;
	
	// Set the ranges
	var	x = d3.time.scale().range([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom");
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).orient("left").ticks(5).orient("left").ticks(5).tickFormat(function(d){
			return "$"+d3.format(",")(d);
		});
	
	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = "<span style='color:white'>"
							+d.tiplabel+" : "+d.value + "</span>";
					});
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart2")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
			.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	
	svg.call(tip);
	//Get the data
	d3.csv("VA_gdp.csv", function(error, data) {
		data.forEach(function(d) {	
			if(d.date==null)
				d.date=0;
			d.label = formatDate(d.date);
			d.tiplabel = d.date;
			d.value = +d[regionNames[region]]/1000;
			
		});
	 
		// Scale the range of the data
		x.domain(d3.extent(data,function(d) {
			return d.label; 
			}));
		y.domain([0, d3.max(data, function(d) {
			
				return d.value;})]);

	  svg.selectAll("bar")
	      .data(data)
	      .enter().append("rect")
	      .style("fill", "#4dc3ff")
	      .attr("x", function(d) { 
	    	
	    	  return x(d.label); })
	      .attr("width",5)
	      .attr("y", function(d) {
	    	  return y(d.value); })
	      .attr("height", function(d) {
	    	 
	    	  return height - y(d.value);})
	    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
	  .attr("class", "x axis")
	  .attr("transform", "translate(0," + height + ")")
	  .call(xAxis)
	.selectAll("text")
	  .style("text-anchor", "end")
	  .attr("dx", "-.8em")
	  .attr("dy", "-.55em")
	  .attr("transform", "rotate(-90)" );

	 svg.append("g")
	  .attr("class", "y axis")
	  .call(yAxis)
	.append("text")
	  .attr("transform", "rotate(-90)")
	  .attr("y", 6)
	  .attr("dy", ".71em")
	  .style("text-anchor", "end")
	  .text(" ");
	 
	});
	
}

function showRevparGraph(region)
{
	var regionNames={"Virginia":"theCommonwealthofVirginia","VirginiaBeach":"HamptonRoadsMarket","Washington":"WashingtonDC-MD-VAMarket","Blacksburg":"Blacksburg/WythevilleMarket","Charlottesville":"CharlottesvilleMarket"
		,"Harrisonburg":"","Lynchburg":"LynchburgMarket","Richmond":"Richmond/PetersburgMarket","Roanoke":"RoanokeMarket","Staunton":"Staunton/HarrisonburgMarket","Harrisonburg":"Staunton/HarrisonburgMarket"};
	
	console.log("Region Key:"+regionNames[region]);
	$("#ceapChart").empty();
	$("#ceapChart").append("<h3 align=\"center\">Hotel Revenue per Available Room by Month</h3>");
	console.log(region);
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	var formatDate = d3.time.format("%b-%Y").parse;
	
	// Set the ranges
	var	x = d3.time.scale().range([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom");
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).tickFormat(function(d){
			return "$"+d3.format(",.2f")(d);
		});
	
	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = " <span style='color:white'>"
							+d.tiplabel+":"+d.value + "</span>";
					});
	    
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
			.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	
	svg.call(tip);
	 
	// Get the data
	d3.csv("Dashboard_REVPAR_Month1.csv", function(error, data) {
		data.forEach(function(d) {
//			if(parseInt(d.Year)>=2005)
//			{
				d.date=d.Year;
				d.label = formatDate(d.Month+"-"+d.Year);
				d.tiplabel = d.Month+"-"+d.Year;
				d.value = +d[regionNames[region]];
//			}
		});
		
	 
		// Scale the range of the data
		x.domain(d3.extent(data,function(d) {
			if(parseInt(d.date)>=2005)
				{
			return d.label; 
			}}));
		y.domain([0, d3.max(data, function(d) {
			if(parseInt(d.date)>=2005)
			{
				return d.value;}})]);
		
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	//if(parseInt(d.Year)>=2005)
		{ console.log(formatDate(d.Month+"-"+d.Year)); return x(formatDate(d.Month+"-"+d.Year)); }})
	    .y(function(d) {// if(parseInt(d.Year)>=2005)
		{ return y(d[regionNames[region]]);} });

//	  svg.selectAll("bar")
//	      .data(data)
//	      .enter().append("rect")
//	       .filter(function(d) { return parseInt(d.date)>=2005 })
//	      .style("fill", "#4dc3ff")
//	      .attr("x", function(d) { 
//	    	  if(parseInt(d.date)>=2005)
//				{
//	    	  return x(d.label);} })
//	      .attr("width",5)
//	      .attr("y", function(d) { 
//	    	  if(parseInt(d.date)>=2005)
//				{
//	    	  return y(d.value); }})
//	      .attr("height", function(d) {
//	    	  if(parseInt(d.date)>=2005)
//				{
//	    	  return height - y(d.value)}})
//	    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
	

	    // Add the scatterplot
	    svg.selectAll("dot")
	        .data(data)
	      .enter().append("circle")
	        .attr("r", 3.5)
	        .attr("cx", function(d) {if(parseInt(d.date)>=2005) {return x(d.label);} })
	        .attr("cy", function(d) {if(parseInt(d.date)>=2005) {return y(d.value);} })
	        .style("fill","#4dc3ff")
	         .on('mouseover', tip.show).on('mouseout', tip.hide);

	  svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," + height + ")")
      .call(xAxis)
    .selectAll("text")
      .style("text-anchor", "end")
      .attr("dx", "-.8em")
      .attr("dy", "-.55em")
      .attr("transform", "rotate(-90)" );

	 svg.append("g")
      .attr("class", "y axis")
      .call(yAxis)
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", ".71em")
      .style("text-anchor", "end")
      .text(" ");
	 
	// Add the valueline path.
	    svg.append("path")
	        .attr("class", "line")
	        .attr("d", valueline(data));
	 
	});
}

function showLabourForce(region)
{
	var regionNames={"Virginia":"VALF","VirginiaBeach":"VIRG251LF","Washington":"WASH911LF","Blacksburg":"BLAC951LF","Charlottesville":"CHAR851LF",
		"Winchester":"WINC051LF","Harrisonburg":"HARR551LF","Lynchburg":"LYNC351LF","Richmond":"RICH051LF","Roanoke":"ROAN251LF","Staunton":"","Harrisonburg":"HARR551LF"};
	var domainStart=0;
	console.log("Region Key:"+regionNames[region]);
	$("#ceapChart").empty();
	if(region=="Staunton")
	{
		showLabourForceNsa(region);
	return;
	}
	
	$("#ceapChart").append("<h3 align=\"center\">Seasonally Adjusted Size of Labor Force by Month</h3><h5 align=\"center\">(In Thousands)</h5>");
	console.log(region);
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	var formatDate = d3.time.format("%m/%d/%Y").parse;
	
	// Set the ranges
	var	x = d3.time.scale().range([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom");
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).tickFormat(function(d){
			return d3.format(",.2f")(d);
		});
	
	var tip;
	var valueline = d3.svg.line()
    .x(function(d) { 
    	
	  return x(d.label); })
    .y(function(d) {
	 return y(d.value); });
	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = " <span style='color:white'>"
							+d.tiplabel+" : "+d.value + "</span>";
					});
		 
	    
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
			.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	
	svg.call(tip);
	 var newdata=[];
	// Get the data
	d3.csv("LABOR_FORCE_lf_ssa_Monthly.csv", function(error, data) {
		data.forEach(function(d) {	
			if(d[regionNames[region]]!=0)
			{
			newdata.push(d);
			d.label = formatDate(d.Date);
			
			d.tiplabel = d.Date;
			d.value = +d[regionNames[region]]/1000;
			}
		});
	 
		// Scale the range of the data
		x.domain(d3.extent(data,function(d) {
			
			return d.label; 
			}));
		
		
		switch(region)
		{
		case "Virginia" : domainStart=3400;
		break;
		case "VirginiaBeach" : domainStart=700;
		break;
		case "Blacksburg" : domainStart=65; 
			break;
		case "Charlottesville" : domainStart=90; 
		break;
		case "Harrisonburg" : domainStart=54; 
		break;
		case "Lynchburg" : domainStart=90; 
		break;
		case "Richmond" : domainStart=540; 
		break;
		case "Roanoke" : domainStart=120; 
		break;
		case "Staunton" : domainStart=40; 
		break;
		case "Winchester" : domainStart=50; 
		break;
		case "Washington" : domainStart=2400; 
		break;
		default : domainStart=0; 
		break;
		
		}
		y.domain([domainStart, d3.max(data, function(d) {
			
			return Math.ceil(d.value);})]);
		
//	  svg.selectAll("bar")
//	      .data(data)
//	      .enter().append("rect")
//	      .style("fill", "#4dc3ff")
//	      .attr("x", function(d) { 
//	    	
//	    	  return x(d.label); })
//	      .attr("width",5)
//	      .attr("y", function(d) { 
//	    	  
//	    	  return y(d.value); })
//	      .attr("height", function(d) {
//	    	 
//	    	  return height - y(d.value);})
//	    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
		
		// Add the valueline path.
		 svg.append("path")
	        .attr("class", "line")
	        .attr("d", valueline(newdata));
	  
	  // Add the scatterplot
	    svg.selectAll("dot")
	        .data(data)
	      .enter().append("circle")
	      .filter(function(d) { return d[regionNames[region]] !=0 })
	        .attr("r", 3.5)
	        .attr("cx", function(d) {return x(d.label); })
	        .attr("cy", function(d) {return y(d.value); })
	        .style("fill","#4dc3ff")
	         .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," + height + ")")
      .call(xAxis)
    .selectAll("text")
      .style("text-anchor", "end")
      .attr("dx", "-.8em")
      .attr("dy", "-.55em")
      .attr("transform", "rotate(-90)" );

	 svg.append("g")
      .attr("class", "y axis")
      .call(yAxis)
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", ".71em")
      .style("text-anchor", "end")
      .text(" ");
	 
	});
	showLabourForceNsa(region);
}

function showFHFA(region)
{
	var regionNames={"Virginia":"VA","VirginiaBeach":"VB-Norfolk","Washington":"Washington-Alexandria","Blacksburg":"Blacksburg","Charlottesville":"Charlottesville",
			"Winchester":"Winchester","Harrisonburg":"Harrisonburg","Lynchburg":"Lynchburg","Richmond":"Richmond","Roanoke":"Roanoke","Staunton":"Staunton-Waynesboro","Harrisonburg":"Harrisonburg"};
		
		console.log("Region Key:"+regionNames[region]);
		$("#ceapChart").empty();
		$("#ceapChart").append("<h3 align=\"center\">FHFA Housing price Index by Quarter</h3>");
		console.log(region);
		var	margin = {top: 10, right: 20, bottom: 40, left: 70},
			width = 900 - margin.left - margin.right,
			height = 500 - margin.top - margin.bottom;
		
		var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5);
		
		var tip;

		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+":"+d.value + "</span>";
						});
			 
		    
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		
		 
		 
		// Get the data
		d3.csv("FHFA_Quarterly.csv", function(error, data) {
			data.forEach(function(d) {	
				
				d.label = formatDate(d.Date);
				
				d.tiplabel = d.Date;
				d.value = +d[regionNames[region]];
			});
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				return d.label; 
				}));
			y.domain([0, d3.max(data, function(d) {
				
					return d.value;})]);

		  svg.selectAll("bar")
		      .data(data)
		      .enter().append("rect")
		      .style("fill", "#4dc3ff")
		      .attr("x", function(d) { 
		    	 
		    	  return x(d.label); })
		      .attr("width",5)
		      .attr("y", function(d) { 
		    	  
		    	  return y(d.value); })
		      .attr("height", function(d) {
		    	 
		    	  return height - y(d.value);})
		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
//		  svg.append("text").attr("x", (width / 2))             
//	        .attr("y", 0 - (margin.top / 2))
//	        .attr("text-anchor", "middle")  
//	        .attr("text-color","black")
//	        .style("font-size", "16px") 
//	        .text("FHFA Housing price Index by Quarter");
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		
		 
		});
		
}
function showWeeklyWages(region)
{
	var regionNames={"Virginia":"Virginia","VirginiaBeach":"Hampton Roads","Washington":"Wash DC","Blacksburg":"Blacksburg","Charlottesville":"Charlottesville",
			"Winchester":"Winchestor","Harrisonburg":"Harrisonburg","Lynchburg":"Lynchburg","Richmond":"Richmond","Roanoke":"Roanoke","Staunton":"Staunton","Harrisonburg":"Harrisonburg"};
		
		console.log("Region Key:"+regionNames[region]);
		$("#ceapChart").empty();
		$("#ceapChart").append("<h3 align=\"center\">Average Weekly Wages by Quarter</h3>");
		console.log(region);
		var	margin = {top: 10, right: 20, bottom: 60, left: 70},
			width = 900 - margin.left - margin.right,
			height = 500 - margin.top - margin.bottom;
		
		//var formatDate = d3.time.format("%Y Q%m").parse;
		
		// Set the ranges
		var	x = d3.scale.ordinal().rangeRoundBands([0, width]);  
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5).tickFormat(function(d){
				return "$"+d3.format(",")(d);
			});
		
		var tip;

		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
			 
		    
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		 
		// Get the data
		d3.csv("WEEKLY_WAGES_Quarterly.csv", function(error, data) {
			data.forEach(function(d) {	
				
				d.label = d.Date;
				d.tiplabel = d.Date;
				d.value = +d[regionNames[region]];
			});
		 
			// Scale the range of the data
			x.domain(data.map(function(d) {
				
				return d.label; 
				}));
			y.domain([0, d3.max(data, function(d) {
				
					return d.value;})]);

		  svg.selectAll("bar")
		      .data(data)
		      .enter().append("rect")
		      .style("fill", "#4dc3ff")
		      .attr("x", function(d) { 
		    	
		    	  return x(d.label); })
		      .attr("width",10)
		      .attr("y", function(d) { 
		    	  
		    	  return y(d.value); })
		      .attr("height", function(d) {
		    	 
		    	  return height - y(d.value);})
		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
}
function showTNFM(region)
{
	var regionNames={"Virginia":"Virginia","VirginiaBeach":"Virginia Beach-Norfolk-Newport News, VA-NC MSA","Washington":"Washington-Arlington-Alexandria, DC-VA-MD-WV MSA, VA part","Blacksburg":"Blacksburg-Christiansburg-Radford, VA MSA","Charlottesville":"Charlottesville, VA MSA",
			"Winchester":"Winchester, VA-WV MSA","Harrisonburg":"Harrisonburg, VA MSA","Lynchburg":"Lynchburg, VA MSA","Richmond":"Richmond, VA MSA","Roanoke":"Roanoke, VA MSA","Staunton":"","Harrisonburg":"Harrisonburg, VA MSA"};
		var domainStart=0;
		console.log("Region Key:"+regionNames[region]);
		$("#ceapChart").empty();
		if(region=="Staunton")
		{
			showTNFMNsa(region);
		return;
		}
		if(region=="Washington")
			$("#ceapChart").append("<h3 align=\"center\">Seasonally Adjusted Nonfarm Employment by Month (VA Portion)</h3> <h5 align=\"center\">(In Thousands)</h5>");		
		else
		$("#ceapChart").append("<h3 align=\"center\">Seasonally Adjusted Nonfarm Employment by Month</h3> <h5 align=\"center\">(In Thousands)</h5>");
		console.log(region);
		var	margin = {top: 10, right: 20, bottom: 40, left: 70},
			width = 900 - margin.left - margin.right,
			height = 500 - margin.top - margin.bottom;
		
		var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5).tickFormat(function(d){
				return d3.format(",.2f")(d);
			});
		
		var tip;
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	
		  return x(d.label); })
	    .y(function(d) {
		 return y(d.value); });
		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
			 
		    
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		var newdata=[];
		// Get the data
		d3.csv("EMPLOYEMENT_TNFALL_emp_Monthly.csv", function(error, data) {
			data.forEach(function(d) {	
				if(d[regionNames[region]]!=0)
				{
				newdata.push(d);
				d.label = formatDate(d.Date);
				d.tiplabel = d.Date;
				d.value = +d[regionNames[region]]/1000;
				}
				});
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				
				return d.label; 
				}));
			switch(region)
			{
			case "Virginia" : domainStart=3500;
			break;
			case "VirginiaBeach" : domainStart=700;
			break;
			case "Blacksburg" : domainStart=65; 
				break;
			case "Charlottesville" : domainStart=90; 
			break;
			case "Harrisonburg" : domainStart=55; 
			break;
			case "Lynchburg" : domainStart=90; 
			break;
			case "Richmond" : domainStart=550; 
			break;
			case "Roanoke" : domainStart=140; 
			break;
			case "Staunton" : domainStart=40; 
			break;
			case "Winchester" : domainStart=50; 
			break;
			case "Washington" : domainStart=2600; 
			break;
			default : domainStart=0; 
			break;
			
			}
			y.domain([domainStart, d3.max(data, function(d) {
				
				return d.value;})]);

//		  svg.selectAll("bar")
//		      .data(data)
//		      .enter().append("rect")
//		      .style("fill", "#4dc3ff")
//		      .attr("x", function(d) { 
//		    	 
//		    	  return x(d.label); })
//		      .attr("width",5)
//		      .attr("y", function(d) { 
//		    	  
//		    	  return y(d.value); })
//		      .attr("height", function(d) {
//		    	 
//		    	  return height - y(d.value);})
//		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
			
			// Add the valueline path.
			 svg.append("path")
		        .attr("class", "line")
		        .attr("d", valueline(newdata));
		  
		  // Add the scatterplot
		    svg.selectAll("dot")
		        .data(data)
		      .enter().append("circle")
		       .filter(function(d) { return d[regionNames[region]] !=0 })
		        .attr("r", 3.5)
		        .attr("cx", function(d) {return x(d.label); })
		        .attr("cy", function(d) {return y(d.value); })
		        .style("fill","#4dc3ff")
		         .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
	showTNFMNsa(region);
}
function showTNFMNsa(region)
{
	var regionNames={"Virginia":"Virginia","VirginiaBeach":"Virginia Beach-Norfolk-Newport News, VA-NC MSA","Washington":"Washington-Arlington-Alexandria, DC-VA-MD-WV MSA, VA part","Blacksburg":"Blacksburg-Christiansburg-Radford, VA MSA","Charlottesville":"Charlottesville, VA MSA",
			"Winchester":"Winchester, VA-WV MSA","Harrisonburg":"Harrisonburg, VA MSA","Lynchburg":"Lynchburg, VA MSA","Richmond":"Richmond, VA MSA","Roanoke":"Roanoke, VA MSA","Staunton":"Staunton-Waynesboro, VA","Harrisonburg":"Harrisonburg, VA MSA"};
		
		console.log("Region Key:"+regionNames[region]);
		$("#ceapChart2").empty();
		var domainStart=0;
		if(region=="Washington")
			$("#ceapChart2").append("<br><br><h3 align=\"center\">Not Seasonally Adjusted Nonfarm Employment by Month (VA Portion)</h3><h5 align=\"center\">(In Thousands)</h5>");
		else
		$("#ceapChart2").append("<br><br><h3 align=\"center\">Not Seasonally Adjusted Nonfarm Employment by Month</h3><h5 align=\"center\">(In Thousands)</h5>");
		console.log(region);
		var	margin = {top: 10, right: 20, bottom: 40, left: 70},
			width = 900 - margin.left - margin.right,
			height = 500 - margin.top - margin.bottom;
		
		var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5).tickFormat(function(d){
				return d3.format(",.2f")(d);
			});
		
		var tip;
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	
		  return x(d.label); })
	    .y(function(d) {
		 return y(d.value); });
		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
			 
		    
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart2")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		 
		var newdata=[];
		// Get the data
		d3.csv("EMPLOYEMENT_TNFALLNSA_emp_Monthly2.csv", function(error, data) {
			data.forEach(function(d) {	
				if(d[regionNames[region]]!=0)
					{
					newdata.push(d);
				d.label = formatDate(d.Date);
				if(d[regionNames[region]]==null)
					d[regionNames[region]]=3000;
				d.tiplabel = d.Date;
				d.value = +d[regionNames[region]]/1000;
				//console.log(d.value+"-"+d.Date);
					}
				
			});
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				if(d[regionNames[region]]!=0) {
					return d.label; 
				}
				
				}));
			switch(region)
			{
			case "Virginia" : domainStart=3500;
			break;
			case "VirginiaBeach" : domainStart=700;
			break;
			case "Blacksburg" : domainStart=65; 
				break;
			case "Charlottesville" : domainStart=90; 
			break;
			case "Harrisonburg" : domainStart=55; 
			break;
			case "Lynchburg" : domainStart=90; 
			break;
			case "Richmond" : domainStart=550; 
			break;
			case "Roanoke" : domainStart=140; 
			break;
			case "Staunton" : domainStart=40; 
			break;
			case "Winchester" : domainStart=50; 
			break;
			case "Washington" : domainStart=2600; 
			break;
			default : domainStart=0; 
			break;
			
			}
			y.domain([domainStart, d3.max(data, function(d) {
				if(d[regionNames[region]]!=0) {
				return d.value;
				}
				})]);

//		  svg.selectAll("bar")
//		      .data(data)
//		      .enter().append("rect")
//		      .style("fill", "#4dc3ff")
//		      .attr("x", function(d) { 
//		    	 
//		    	  return x(d.label); })
//		      .attr("width",5)
//		      .attr("y", function(d) { 
//		    	  
//		    	  return y(d.value); })
//		      .attr("height", function(d) {
//		    	 
//		    	  return height - y(d.value);})
//		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
			
			// Add the valueline path.
			 svg.append("path")
		        .attr("class", "line")
		        .attr("d", valueline(newdata));
		  
		  // Add the scatterplot
		    svg.selectAll("dot")
		        .data(data)
		        .enter().append("circle")
		        .filter(function(d) { return d[regionNames[region]] !=0 })
		        .attr("r", 3.5)
		        .attr("cx", function(d) {
		        	return x(d.label); })
		        .attr("cy", function(d) {return y(d.value); })
		        .style("fill","#4dc3ff")
		         .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
		
		if(region=="Washington")
			showWashSSA(region);
}
function showWashSSA(region)
{
	
	$("#ceapChart3").empty();
	$("#ceapChart3").append("<br><h3 align=\"center\">Seasonally Adjusted Nonfarm Employment by Month</h3> <h5 align=\"center\">(In Thousands)</h5>");
	console.log(region);
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	var formatDate = d3.time.format("%m/%d/%Y").parse;
	
	// Set the ranges
	var	x = d3.time.scale().range([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom");
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).tickFormat(function(d){
			return d3.format(",.2f")(d);
		});
	
	var tip;
	var valueline = d3.svg.line()
    .x(function(d) { 
    	
	  return x(d.label); })
    .y(function(d) {
	 return y(d.value); });
	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = " <span style='color:white'>"
							+d.tiplabel+" : "+d.value + "</span>";
					});
		 
	    
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart3")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
			.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	
	svg.call(tip);
	 
	// Get the data
	d3.csv("Washington dc MSA Employment latest.csv", function(error, data) {
		data.forEach(function(d) {	
			d.label = formatDate(d.date);
			d.tiplabel = d.date;
			d.value = +d["Ssa_value"];
			//console.log(d.value+"-"+d[regionNames[region]]+"-"+d.Date);
		});
	 
		// Scale the range of the data
		x.domain(d3.extent(data,function(d) {
			return d.label; 
			}));
		
		y.domain([2400, d3.max(data, function(d) {
			
			return Math.ceil(d.value);})]);

		 svg.append("path")
	        .attr("class", "line")
	        .attr("d", valueline(data));
	  
	  // Add the scatterplot
	    svg.selectAll("dot")
	        .data(data)
	      .enter().append("circle")
	        .attr("r", 3.5)
	        .attr("cx", function(d) {return x(d.label); })
	        .attr("cy", function(d) {return y(d.value); })
	        .style("fill","#4dc3ff")
	         .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," + height + ")")
      .call(xAxis)
    .selectAll("text")
      .style("text-anchor", "end")
      .attr("dx", "-.8em")
      .attr("dy", "-.55em")
      .attr("transform", "rotate(-90)" );

	 svg.append("g")
      .attr("class", "y axis")
      .call(yAxis)
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", ".71em")
      .style("text-anchor", "end")
      .text(" ");
	 
	});
	showWashNSA(region);
}
function showWashNSA(region)
{
	
	$("#ceapChart4").empty();
	$("#ceapChart4").append("<br><h3 align=\"center\">Not Seasonally Adjusted Nonfarm Employment by Month</h3> <h5 align=\"center\">(In Thousands)</h5>");
	console.log(region);
	var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
	
	var formatDate = d3.time.format("%m/%d/%Y").parse;
	
	// Set the ranges
	var	x = d3.time.scale().range([0, width]);
	var	y = d3.scale.linear().range([height, 0]);
	 
	// Define the axes
	var	xAxis = d3.svg.axis().scale(x)
		.orient("bottom");
	 
	var	yAxis = d3.svg.axis().scale(y)
		.orient("left").ticks(5).tickFormat(function(d){
			return d3.format(",.2f")(d);
		});
	
	var tip;
	var valueline = d3.svg.line()
    .x(function(d) { 
    	
	  return x(d.label); })
    .y(function(d) {
	 return y(d.value); });
	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = " <span style='color:white'>"
							+d.tiplabel+" : "+d.value + "</span>";
					});
		 
	    
	// Adds the svg canvas
	var	svg = d3.select("#ceapChart4")
		.append("svg")
			.attr("width", width + margin.left + margin.right)
			.attr("height", height + margin.top + margin.bottom)
			.append("g")
			.attr("transform", "translate(" + margin.left + ","+5+")");
	
	svg.call(tip);
	 
	// Get the data
	d3.csv("Washington dc MSA Employment latest.csv", function(error, data) {
		data.forEach(function(d) {	
			d.label = formatDate(d.date);
			d.tiplabel = d.date;
			d.value = +d["Nsa_value"];
			//console.log(d.value+"-"+d[regionNames[region]]+"-"+d.Date);
		});
	 
		// Scale the range of the data
		x.domain(d3.extent(data,function(d) {
			return d.label; 
			}));
		
		y.domain([2400, d3.max(data, function(d) {
			
			return Math.ceil(d.value);})]);

		 svg.append("path")
	        .attr("class", "line")
	        .attr("d", valueline(data));
	  
	  // Add the scatterplot
	    svg.selectAll("dot")
	        .data(data)
	      .enter().append("circle")
	      
	        .attr("r", 3.5)
	        .attr("cx", function(d) {return x(d.label); })
	        .attr("cy", function(d) {return y(d.value); })
	        .style("fill","#4dc3ff")
	         .on('mouseover', tip.show).on('mouseout', tip.hide);
	  
	  svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," + height + ")")
      .call(xAxis)
    .selectAll("text")
      .style("text-anchor", "end")
      .attr("dx", "-.8em")
      .attr("dy", "-.55em")
      .attr("transform", "rotate(-90)" );

	 svg.append("g")
      .attr("class", "y axis")
      .call(yAxis)
    .append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 6)
      .attr("dy", ".71em")
      .style("text-anchor", "end")
      .text(" ");
	 
	});
	
}
function showLabourForceNsa(region)
{
	var regionNames={"Virginia":"VALFN","VirginiaBeach":"VIRG251LFN","Washington":"WASH911LFN","Blacksburg":"BLAC951LFN","Charlottesville":"CHAR851LFN",
			"Winchester":"WINC051LFN","Harrisonburg":"HARR551LFN","Lynchburg":"LYNC351LFN","Richmond":"RICH051LFN","Roanoke":"ROAN251LFN","Staunton":"staunton","Harrisonburg":"HARR551LFN"};
	var domainStart=0;
		console.log("Region Key:"+regionNames[region]);
		$("#ceapChart2").empty();
		$("#ceapChart2").append("<br><br><h3 align=\"center\">Not Seasonally Adjusted Size of Labor Force by Month</h3><h5 align=\"center\">(In Thousands)</h5>");
		console.log(region);
		var	margin = {top: 10, right: 20, bottom: 40, left: 70},
			width = 900 - margin.left - margin.right,
			height = 500 - margin.top - margin.bottom;
		
		var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5).tickFormat(function(d){
				return d3.format(",.2f")(d);
			});
		
		var tip;
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	
		  return x(d.label); })
	    .y(function(d) {
		 return y(d.value); });
		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
			 
		    
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart2")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		 var newdata=[];
		// Get the data
		d3.csv("LABOR_FORCE_lf_nsa_Monthly2.csv", function(error, data) {
			data.forEach(function(d) {	
				if(d[regionNames[region]]!=0)
				{
				newdata.push(d);
				d.label = formatDate(d.Date);
				d.tiplabel = d.Date;
				d.value = +d[regionNames[region]]/1000;
				//console.log(d.value+"-"+d[regionNames[region]]+"-"+d.Date);
				}
			});
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				return d.label; 
				}));
			switch(region)
			{
			case "Virginia" : domainStart=3400;
			break;
			case "VirginiaBeach" : domainStart=700;
			break;
			case "Blacksburg" : domainStart=65; 
				break;
			case "Charlottesville" : domainStart=90; 
			break;
			case "Harrisonburg" : domainStart=54; 
			break;
			case "Lynchburg" : domainStart=90; 
			break;
			case "Richmond" : domainStart=540; 
			break;
			case "Roanoke" : domainStart=120; 
			break;
			case "Staunton" : domainStart=40; 
			break;
			case "Winchester" : domainStart=50; 
			break;
			case "Washington" : domainStart=2400; 
			break;
			default : domainStart=0; 
			break;
			
			}
			y.domain([domainStart, d3.max(data, function(d) {
				
				return Math.ceil(d.value);})]);

//		  svg.selectAll("bar")
//		      .data(data)
//		      .enter().append("rect")
//		      .style("fill", "#4dc3ff")
//		      .attr("x", function(d) { 
//		    	 
//		    	  return x(d.label); })
//		      .attr("width",5)
//		      .attr("y", function(d) {
//		    	  return y(d.value); })
//		      .attr("height", function(d) {
//		    	 
//		    	  return height - y(d.value);})
//		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
			// Add the valueline path.
			 svg.append("path")
		        .attr("class", "line")
		        .attr("d", valueline(newdata));
		  
		  // Add the scatterplot
		    svg.selectAll("dot")
		        .data(data)
		      .enter().append("circle")
		      .filter(function(d) { return d[regionNames[region]] !=0 })
		        .attr("r", 3.5)
		        .attr("cx", function(d) {return x(d.label); })
		        .attr("cy", function(d) {return y(d.value); })
		        .style("fill","#4dc3ff")
		         .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
}

function showUiClaims(region)
{
	var regionNames={"Virginia":"Virginia","VirginiaBeach":"Virginia Beach-Norfolk-Newport News, VA-NC Metropolitan Statistical Area, VA part","Washington":"Washington-Arlington-Alexandria, DC-VA-MD-WV Metropolitan Statistical Area, VA part","Blacksburg":"Blacksburg-Christiansburg-Radford, VA Metropolitan Statistical Area","Charlottesville":"Charlottesville, VA Metropolitan Statistical Area",
			"Winchester":"Winchester, VA-WV Metropolitan Statistical Area, VA part","Harrisonburg":"Harrisonburg, VA Metropolitan Statistical Area","Lynchburg":"Lynchburg, VA Metropolitan Statistical Area","Richmond":"Richmond, VA Metropolitan Statistical Area","Roanoke":"Roanoke, VA Metropolitan Statistical Area","Staunton":"Staunton-Waynesboro, VA Metropolitan Statistical Area","Harrisonburg":"Harrisonburg, VA Metropolitan Statistical Area"};
		
		console.log("Region Key:"+regionNames[region]);
		$("#ceapChart").empty();
		$("#ceapChart").append("<h3 align=\"center\">Initial Unemployment Claims by Month</h3>");
		var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
		
		var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5);
		
		var tip;
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	
		  return x(d.label); })
	    .y(function(d) {
		 return y(d.value); });
		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		var newdata=[];
		// Get the data
		d3.csv("uiclaims.csv", function(error, data) {
			data.forEach(function(d) {	
				if(d[regionNames[region]]!=0)
				{
				newdata.push(d);
				d.label = formatDate(d.date);
				d.tiplabel = d.date;
				d.value = +d[regionNames[region]];
				console.log("value="+d.label);
				}
			});
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				return d.label; 
				}));
			y.domain([0, d3.max(data, function(d) {
				
					return d.value;})]);

//		  svg.selectAll("bar")
//		      .data(data)
//		      .enter().append("rect")
//		      .style("fill", "#4dc3ff")
//		      .attr("x", function(d) { 
//		    	 
//		    	  return x(d.label); })
//		      .attr("width",5)
//		      .attr("y", function(d) {
//		    	  return y(d.value); })
//		      .attr("height", function(d) {
//		    	 
//		    	  return height - y(d.value);})
//		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
			
			// Add the valueline path.
			 svg.append("path")
		        .attr("class", "line")
		        .attr("d", valueline(newdata));
		  
		  // Add the scatterplot
		    svg.selectAll("dot")
		        .data(data)
		      .enter().append("circle")
		      .filter(function(d) { return d[regionNames[region]] !=0 })
		        .attr("r", 3.5)
		        .attr("cx", function(d) {return x(d.label); })
		        .attr("cy", function(d) {return y(d.value); })
		        .style("fill","#4dc3ff")
		         .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
		
}

function showTaxRetailSales(region)
{
	var regionNames={"Virginia":"Virginia","VirginiaBeach":"Virginia Beach-Norfolk-Newport News, VA-NC Metropolitan Statistical Area, VA part","Washington":"Washington-Arlington-Alexandria, DC-VA-MD-WV Metropolitan Statistical Area, VA part","Blacksburg":"Blacksburg-Christiansburg-Radford, VA Metropolitan Statistical Area","Charlottesville":"Charlottesville, VA Metropolitan Statistical Area",
			"Winchester":"Winchester, VA-WV Metropolitan Statistical Area, VA part","Harrisonburg":"Harrisonburg, VA Metropolitan Statistical Area","Lynchburg":"Lynchburg, VA Metropolitan Statistical Area","Richmond":"Richmond, VA Metropolitan Statistical Area","Roanoke":"Roanoke, VA Metropolitan Statistical Area","Staunton":"Staunton-Waynesboro, VA Metropolitan Statistical Area","Harrisonburg":"Harrisonburg, VA Metropolitan Statistical Area"};
		
		console.log("Region Key:"+regionNames[region]);
		
		$("#ceapChart").empty();
		$("#ceapChart").append("<h3 align=\"center\">Taxable Retail Sales by Month</h3><h5 align=\"center\">(In Millions of Dollars)</h5>");
		var	margin = {top: 10, right: 20, bottom: 40, left: 90},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
		
		var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5).tickFormat(function(d){
							
							return "$"+d3.format(",.2f")(d);
			});
		
		var tip;
		var domainStart;
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	
		  return x(d.label); })
	    .y(function(d) {
		 return y(d.value); });

		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		var newdata=[];
		// Get the data
		d3.csv("local_option_sales_data.csv", function(error, data) {
			data.forEach(function(d) {	
				if(d[regionNames[region]]!=0)
				{
				newdata.push(d);
				d.label = formatDate(d.date);
				d.tiplabel = d.date;
				d.value = +d[regionNames[region]]/10000;
				}
			});
			switch(region)
			{
			case "Virginia" : domainStart=4000;
			break;
			case "VirginiaBeach" : domainStart=1000;
			break;
			case "Blacksburg" : domainStart=80; 
				break;
			case "Charlottesville" : domainStart=100; 
			break;
			case "Harrisonburg" : domainStart=80; 
			break;
			case "Lynchburg" : domainStart=100; 
			break;
			case "Richmond" : domainStart=800; 
			break;
			case "Roanoke" : domainStart=200; 
			break;
			case "Staunton" : domainStart=60; 
			break;
			case "Winchester" : domainStart=50; 
			break;
			case "Washington" : domainStart=2000; 
			break;
			default : domainStart=0; 
			break;
			
			}
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				return d.label; 
				}));
			y.domain([domainStart, d3.max(data, function(d) {
				
					return d.value;})]);
//
//		  svg.selectAll("bar")
//		      .data(data)
//		      .enter().append("rect")
//		      .style("fill", "#4dc3ff")
//		      .attr("x", function(d) { 
//		    	 
//		    	  return x(d.label); })
//		      .attr("width",5)
//		      .attr("y", function(d) {
//		    	  return y(d.value); })
//		      .attr("height", function(d) {
//		    	 
//		    	  return height - y(d.value);})
//		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
			// Add the valueline path.
			 svg.append("path")
		        .attr("class", "line")
		        .attr("d", valueline(newdata));
		  
		  // Add the scatterplot
		    svg.selectAll("dot")
		        .data(data)
		      .enter().append("circle")
		      .filter(function(d) { return d[regionNames[region]] !=0 })
		        .attr("r", 3.5)
		        .attr("cx", function(d) {return x(d.label); })
		        .attr("cy", function(d) {return y(d.value); })
		        .style("fill","#4dc3ff")
		         .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
		
}

function showUnemploymentRate(region)
{
	var regionNames={"Virginia":"VAUR","VirginiaBeach":"VIRG251UR","Washington":"WASH911UR","Blacksburg":"BLAC951UR","Charlottesville":"CHAR851UR",
			"Winchester":"WINC051UR","Harrisonburg":"HARR551UR","Lynchburg":"LYNC351UR","Richmond":"RICH051UR","Roanoke":"ROAN251UR","Staunton":"","Harrisonburg":"HARR551UR"};
		
		console.log("Region Key:"+regionNames[region]);
		
		
		$("#ceapChart").empty();
		if(region=="Staunton")
			{
			showUnemploymentRateNsa(region);
			return;
			}
		$("#ceapChart").append("<h3 align=\"center\">Seasonally Adjusted Unemployment Rate by Month</h3>");
		//$("#ceapChart").append("<h4 align=\"center\">Seasonally Adjusted</h4>");
		var	margin = {top: 30, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
var formatDate = d3.time.format("%m/%d/%Y").parse;
		
		// Set the ranges
		var	x = d3.time.scale().range([0, width]);
		var	y = d3.scale.linear().range([height, 0]);
		 
		// Define the axes
		var	xAxis = d3.svg.axis().scale(x)
			.orient("bottom");
		 
		var	yAxis = d3.svg.axis().scale(y)
			.orient("left").ticks(5).tickFormat(function(d){
				return parseFloat(d).toFixed(2)+"%";
			});
		
		var tip;
		var valueline = d3.svg.line()
	    .x(function(d) { 
	    	
		  return x(d.label); })
	    .y(function(d) {
		 return y(d.value); });
		tip = d3
				.tip()
				.attr('class', 'd3-tip')
				.offset([ -10, 0 ])
				.html(
						function(d) {
							return value = " <span style='color:white'>"
								+d.tiplabel+" : "+d.value + "</span>";
						});
		// Adds the svg canvas
		var	svg = d3.select("#ceapChart")
			.append("svg")
				.attr("width", width + margin.left + margin.right)
				.attr("height", height + margin.top + margin.bottom)
				.append("g")
				.attr("transform", "translate(" + margin.left + ","+5+")");
		
		svg.call(tip);
		var newdata=[];
		// Get the data
		d3.csv("labor_force_data.csv", function(error, data) {
			data.forEach(function(d) {	
				if(d[regionNames[region]]!=0)
				{
				newdata.push(d);
				d.label = formatDate(d.date);
				d.tiplabel = d.date;
				if(d[regionNames[region]]==null)
					d[regionNames[region]]=0;
				d.value = +d[regionNames[region]];
				}
				
			});
		 
			// Scale the range of the data
			x.domain(d3.extent(data,function(d) {
				return d.label; 
				}));
			y.domain([2,10]);

//		  svg.selectAll("bar")
//		      .data(data)
//		      .enter().append("rect")
//		      .style("fill", "#4dc3ff")
//		      .attr("x", function(d) { 
//		    	 
//		    	  return x(d.label); })
//		      .attr("width",5)
//		      .attr("y", function(d) {
//		    	  return y(d.value); })
//		      .attr("height", function(d) {
//		    	 
//		    	  return height - y(d.value);})
//		    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
			// Add the valueline path.
			 svg.append("path")
		        .attr("class", "line")
		        .attr("d", valueline(newdata));
		  
		  // Add the scatterplot
		    svg.selectAll("dot")
		        .data(data)
		      .enter().append("circle")
		      .filter(function(d) { return d[regionNames[region]] !=0 })
		        .attr("r", 3.5)
		        .attr("cx", function(d) {return x(d.label); })
		        .attr("cy", function(d) {return y(d.value); })
		        .style("fill","#4dc3ff")
		         .on('mouseover', tip.show).on('mouseout', tip.hide);
		  
		  svg.append("g")
	      .attr("class", "x axis")
	      .attr("transform", "translate(0," + height + ")")
	      .call(xAxis)
	      
	    .selectAll("text")
	      .style("text-anchor", "end")
	      .attr("dx", "-.8em")
	      .attr("dy", "-.55em")
	      .attr("transform", "rotate(-90)" );

		 svg.append("g")
	      .attr("class", "y axis")
	      .call(yAxis)
	    .append("text")
	      .attr("transform", "rotate(-90)")
	      .attr("y", 6)
	      .attr("dy", ".71em")
	      .style("text-anchor", "end")
	      .text(" ");
		 
		});
		showUnemploymentRateNsa(region);
}
function showUnemploymentRateNsa(region)
{
	var regionNames={"Virginia":"VAURN","VirginiaBeach":"VIRG251URN","Washington":"WASH911URN","Blacksburg":"BLAC951URN","Charlottesville":"CHAR851URN",
			"Winchester":"WINC051URN","Harrisonburg":"HARR551URN","Lynchburg":"LYNC351URN","Richmond":"RICH051URN","Roanoke":"ROAN251URN","Staunton":"staunton","Harrisonburg":"HARR551URN"};
		
		console.log("Region Key:"+regionNames[region]);
		
		$("#ceapChart2").empty();
		$("#ceapChart2").append("<br><br><h3 align=\"center\">Not Seasonally Adjusted Unemployment Rate by Month</h3>");
		var	margin = {top: 10, right: 20, bottom: 40, left: 70},
		width = 900 - margin.left - margin.right,
		height = 500 - margin.top - margin.bottom;
var formatDate = d3.time.format("%m/%d/%Y").parse;
//Set the ranges
var	x = d3.time.scale().range([0, width]);
var	y = d3.scale.linear().range([height, 0]);
 
// Define the axes
var	xAxis = d3.svg.axis().scale(x)
	.orient("bottom");
 
var	yAxis = d3.svg.axis().scale(y)
	.orient("left").ticks(5).tickFormat(function(d){
		return parseFloat(d).toFixed(2)+"%";
	});

var tip;
var valueline = d3.svg.line()
.x(function(d) { 
	
  return x(d.label); })
.y(function(d) {
 return y(d.value); });
tip = d3
		.tip()
		.attr('class', 'd3-tip')
		.offset([ -10, 0 ])
		.html(
				function(d) {
					return value = "<span style='color:white'>"
						+d.tiplabel+" : "+d.value + "</span>";
				});
// Adds the svg canvas
var	svg = d3.select("#ceapChart2")
	.append("svg")
		.attr("width", width + margin.left + margin.right)
		.attr("height", height + margin.top + margin.bottom)
		.append("g")
		.attr("transform", "translate(" + margin.left + ","+5+")");

svg.call(tip);
var newdata=[];
//Get the data
d3.csv("labor_force_dataNsa.csv", function(error, data) {
	data.forEach(function(d) {	
		if(d[regionNames[region]]!=0)
		{
		newdata.push(d);
		d.label = formatDate(d.date);
		d.tiplabel = d.date;
		if(d[regionNames[region]]==null)
			d[regionNames[region]]=2;
		d.value = +d[regionNames[region]];
		//console.log(d.value);
		}
	});
 
	// Scale the range of the data
	x.domain(d3.extent(data,function(d) {
		return d.label; 
		}));
	y.domain([2,10]);

//  svg.selectAll("bar")
//      .data(data)
//      .enter().append("rect")
//      .style("fill", "#4dc3ff")
//      .attr("x", function(d) { 
//    	 
//    	  return x(d.label); })
//      .attr("width",5)
//      .attr("y", function(d) {
//    	  return y(d.value); })
//      .attr("height", function(d) {
//    	 
//    	  return height - y(d.value);})
//    	   .on('mouseover', tip.show).on('mouseout', tip.hide);
	
	// Add the valueline path.
	 svg.append("path")
       .attr("class", "line")
       .attr("d", valueline(newdata));
 
 // Add the scatterplot
   svg.selectAll("dot")
       .data(data)
     .enter().append("circle")
     .filter(function(d) { return d[regionNames[region]] !=0 })
       .attr("r", 3.5)
       .attr("cx", function(d) {return x(d.label); })
       .attr("cy", function(d) {return y(d.value); })
       .style("fill","#4dc3ff")
        .on('mouseover', tip.show).on('mouseout', tip.hide);
  
  svg.append("g")
  .attr("class", "x axis")
  .attr("transform", "translate(0," + height + ")")
  .call(xAxis)
.selectAll("text")
  .style("text-anchor", "end")
  .attr("dx", "-.8em")
  .attr("dy", "-.55em")
  .attr("transform", "rotate(-90)" );

 svg.append("g")
  .attr("class", "y axis")
  .call(yAxis)
.append("text")
  .attr("transform", "rotate(-90)")
  .attr("y", 6)
  .attr("dy", ".71em")
  .style("text-anchor", "end")
  .text(" ");
 
});
}
