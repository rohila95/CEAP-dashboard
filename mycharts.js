function getGraphValue(path, d) {
	var value = 0;
	switch (path) {
	case 1:
		value = d.Outbound;
		break;
	case 2:
		value = d.Inbound;
		break;
	case 3:
		value = d.TotalLoaded;
		break;
	case 4:
		value = d.EmptyTEU;
		break;
	case 5:
		value = d.TotalTEU;
		break;
	case 6:
		value = d.TotalContainer;
		break;
	case 7:
		value = d.ContainerCargo;
		break;
	case 8:
		value = d.BreakBulkTonnage;
		break;
	case 9:
		value = d.TotalGeneralCargo;
		break;
	case 10:
		value = d.TotalShipCalls;
		break;
	}
	return parseFloat(value);
}

function drawBarChartPOV(path) {
	var margin = {
		top : 60,
		right : 20,
		bottom : 50,
		left : 40,
		padding : 20
	}, width = 600 - margin.left - margin.right, height = 500 - margin.top
			- margin.bottom, padding = 20;

	var x = d3.scale.ordinal().rangeRoundBands([ 0, width ]);

	var y = d3.scale.linear().range([ height, 0 ]);

	var xAxis = d3.svg.axis().scale(x).orient("bottom").ticks(0);

	var yAxis = d3.svg.axis().scale(y).orient("left").ticks(6).tickSize(5);

	var tip;

	tip = d3
			.tip()
			.attr('class', 'd3-tip')
			.offset([ -10, 0 ])
			.html(
					function(d) {
						return value = "<strong>Value:</strong> <span style='color:red'>"
								+ getGraphValue(path, d) + "</span>";
					})

	var svg = d3.select("#chartHolder").append("svg").attr("width",
			width + margin.left + margin.right).attr("height",
			height + margin.top + margin.bottom).append("g").attr("transform",
			"translate(" + margin.left + "," + margin.top + ")");

	svg.call(tip);

	d3.csv("Port of Virginia Statistics_Graph.csv", function(error, data) {
		x.domain(data.map(function(d) {
			if (getGraphValue(path, d) > 0 && getGraphValue(path, d) != null) {
				return (d.Year);
			}
		}));
		y.domain([ 0, d3.max(data, function(d) {
			// console.log("IN csv read:"+getGraphValue(path, d) );
			if (getGraphValue(path, d) > 0) {
				return getGraphValue(path, d);
			}
		}) ]);

		svg.append("g").attr("class", "x axis").attr("transform",
				"translate(0," + (height) + ")").call(xAxis).selectAll("text")
				.attr("y", -15).attr("x", -25).attr("dy", ".35em").attr(
						"transform", "rotate(270)").style("text-anchor",
						"middle");

		svg.append("g").attr("class", "y axis").call(yAxis).append("text")
				.attr("transform", "rotate(-90)").attr("y", 0).attr("dy",
						".35em").style("text-anchor", "end").text("");

		svg.selectAll(".bar").data(data).enter().append("rect").attr("class",
				"bar").attr("x", function(d) {
			if (getGraphValue(path, d) > 0) {
				return x(d.Year);
			}
		}).attr("width", x.rangeBand()).attr("y", function(d) {
			// console.log(path+":"+d);
			if (getGraphValue(path, d) > 0) {
				return y(getGraphValue(path, d));
			}
		}).attr("height", function(d) {
			if (getGraphValue(path, d) > 0) {
				return height - y(getGraphValue(path, d));
			}
		})
		// .filter("display", function(d) {return d ==
		// null;}).remove()
		.on('mouseover', tip.show).on('mouseout', tip.hide)

	});
}

function getPieGraphValue(path, d) {
	var value = 0;
	switch (parseInt(path)) {
	case 1991:
		value = d.Y1991;
		break;
	case 1992:
		value = d.Y1992;
		break;
	case 1993:
		value = d.Y1993;
		break;
	case 1994:
		value = d.Y1994;
		break;
	case 1995:
		value = d.Y1995;
		break;
	case 1996:
		value = d.Y1996;
		break;
	case 1997:
		value = d.Y1997;
		break;
	case 1998:
		value = d.Y1998;
		break;
	case 1999:
		value = d.Y1999;
		break;
	case 2000:
		value = d.Y2000;
		break;
	case 2001:
		value = d.Y2001;
		break;
	case 2002:
		value = d.Y2002;
		break;
	case 2003:
		value = d.Y2003;
		break;
	case 2004:
		value = d.Y2004;
		break;
	case 2005:
		value = d.Y2005;
		break;
	case 2006:
		value = d.Y2006;
		break;
	case 2007:
		value = d.Y2007;
		break;
	case 2008:
		value = d.Y2008;
		break;
	case 2009:
		value = d.Y2009;
		break;
	case 2010:
		value = d.Y2010;
		break;
	case 2011:
		value = d.Y2011;
		break;
	case 2012:
		value = d.Y2012;
		break;
	case 2013:
		value = d.Y2013;
		break;
	case 2014:
		value = d.Y2014;
		break;
	case 2015:
		value = d.Y2015;
		break;
	case 2016:
		value = d.Y2016;
		break;
	}
	return parseFloat(value);
}

function inboundTEU(x) {

	var width = 300, height = 300, radius = Math.min(width, height) / 3;

	var color = d3.scale.ordinal().range(
			[ "#ff4d4d", "#66ccff", "#ffff99", "#ff00ff", "#ff9933", "#00ffcc",
					"#993333", "#2ca02c" ]);

	var arc = d3.svg.arc().outerRadius(radius - 10).innerRadius(0);

	var labelArc = d3.svg.arc().outerRadius(radius + 5).innerRadius(radius + 5);

	var pie = d3.layout.pie().sort(null).value(function(d) {
		return getPieGraphValue(x, d);
	});

	var svg = d3.select("#chartTEUSFirst").append("svg").attr("width", width)
			.attr("height", height).append("g").attr("transform",
					"translate(" + width / 2 + "," + height / 2 + ")");

	d3.csv("TEUS MONTHLY PORT LATEST_Graph.csv", function(error, data) {
		if (error)
			throw error;

		var g = svg.selectAll(".arc").data(pie(data)).enter().append("g").attr(
				"class", "arc");

		g.append("path").attr("d", arc).style("fill", function(d) {
			return color(d.data.Region);
		});

		g.append("text").attr("transform", function(d) {
			return "translate(" + labelArc.centroid(d) + ")";
		}).attr("dy", ".10em")
				.text(
						function(d) {
							// console.log("pei:" + getPieGraphValue(x,
							// d.data));
							var value = Math.round(10000 * (getPieGraphValue(x,
									d.data))) / 100;
							return value;
						});

		var legend = d3.select("#chartTEUSLegend").append("svg").attr("class",
				"legend").attr("width", radius * 2).attr("height", radius * 2)
				.selectAll("g").data(data).enter().append("g").attr(
						"transform", function(data, i) {
							return "translate(0," + i * 20 + ")";
						});

		legend.append("rect").attr("width", 18).attr("height", 18).style(
				"fill", function(data, i) {
					return color(i);
				});

		legend.append("text").attr("x", 24).attr("y", 9).attr("dy", ".35em")
				.text(function(data) {
					return data.Region;
				});
	});

	var pie1 = d3.layout.pie().sort(null).value(function(d) {
		return d.Y2015;
	});

	var svg1 = d3.select("#chartTEUSSecond").append("svg").attr("width", width)
			.attr("height", height).append("g").attr("transform",
					"translate(" + width / 2 + "," + height / 2 + ")");

	d3.csv("TEUS MONTHLY PORT LATEST_Graph.csv", function(error, data) {
		if (error)
			throw error;

		var g = svg1.selectAll(".arc").data(pie1(data)).enter().append("g")
				.attr("class", "arc");

		g.append("path").attr("d", arc).style("fill", function(d) {
			return color(d.data.Region);
		});

		g.append("text").attr("transform", function(d) {
			return "translate(" + labelArc.centroid(d) + ")";
		}).attr("dy", ".10em").text(function(d) {
			return Math.round(10000 * (d.data.Y2015)) / 100;
		});
	});
}

function getHotelPieChartRegion(x,d, region) {
	var hamptonRoadRegions = [ "Norfolk/Portsmouth Market",
			"Chesapeake/Suffolk Market", "Williamsburg Market",
			"Virginia Beach Market", "Newport News/Hampton Market" ];
	
	var virginiaRegions=["Richmond/Petersburg Market","Staunton/Harrisonburg Market","Virginia Portion of Washington DC",
	                     "Blacksburg/Wytheville Market","Charlottesville Market","Lynchburg Market","Roanoke Market","Other Markets","Hampton Roads Market","Virginia Portion of Bristol/Kingsport"];
	var dcRegions=["DC CBD WMS","Maryland Portion WMS","Remaining WMS","Virginia Portion of Washington DC WMS"];
	var dcRoomRegions=["DC CBD ARS","Maryland Portion ARS","Remaining ARS","Virginia Portion of Washington DC ARS"];
	

	if (region == 1 && $.inArray(d.Region, hamptonRoadRegions) >= 0) {
			
			return d["Y"+x];
			//return getPieGraphValue(x, d);
	}
	else if(region == 2 && $.inArray(d.Region, virginiaRegions) >= 0)
		{
		return d["Y"+x];
		}
	else if(region == 3 && $.inArray(d.Region, dcRegions) >= 0)
	{
		return d["Y"+x];
	}
	else if(region == 4 && $.inArray(d.Region, dcRoomRegions) >= 0)
	{
		return d["Y"+x];
	}
	else
		{
		return 0;
		}
}

function drawHotelPieChart(x, region) {

	var hamptonRoadRegions = [ "Norfolk/Portsmouth Market",
			"Chesapeake/Suffolk Market", "Williamsburg Market",
			"Virginia Beach Market", "Newport News/Hampton Market" ];
	
	var virginiaRegions=["Richmond/Petersburg Market","Staunton/Harrisonburg Market","Virginia Portion of Washington DC",
	                     "Blacksburg/Wytheville Market","Charlottesville Market","Lynchburg Market","Roanoke Market","Other Markets","Hampton Roads Market","Virginia Portion of Bristol/Kingsport"];
	var dcRegions=["DC CBD WMS","Maryland Portion WMS","Remaining WMS","Virginia Portion of Washington DC WMS"];
	var dcRoomRegions=["DC CBD ARS","Maryland Portion ARS","Remaining ARS","Virginia Portion of Washington DC ARS"];

	var width = 300, height = 300, radius = Math.min(width, height) / 3;

	var color = d3.scale.category20();
	var color1 = d3.scale.category20();

	var arc = d3.svg.arc().outerRadius(radius - 10).innerRadius(0);

	var labelArc = d3.svg.arc().outerRadius(radius + 5).innerRadius(radius + 5);

	var pie = d3.layout.pie().sort(null).value(function(d) {
		return getHotelPieChartRegion(x,d,region);
	});
	
	//console.log(pie);

	var svg = d3.select("#chartHotelPieFirst").append("svg").attr("width",
			width).attr("height", height).append("g").attr("transform",
			"translate(" + width / 2 + "," + height / 2 + ")");
	
	svg.append("text")
    .attr("x", 0)             
    .attr("y", -130)
    .attr("text-anchor", "middle")  
    .style("font-size", "12px") 
    .style("text-decoration", "underline")
    .style("font-weight", "bold")  
    .text(x);

	d3.csv("HotelData_PieGraph.csv", function(error, data) {
		if (error)
			{
			console.log("error");
			throw error;
			}
		
		//console.log(data);
			

		var g = svg.selectAll(".arc").data(
				pie(data).filter(
						function(d) {
							//console.log("d is =="+d.data.Region);
							//console.log(region+":"+d.data.Region);
							if (region == 1 && $.inArray(d.data.Region,hamptonRoadRegions) >= 0) {
								//console.log("1");
								return true;
							}
							else if(region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0)
							{
								//console.log("2");
								return true;
							}
							else if(region == 3 && $.inArray(d.data.Region, dcRegions) >= 0)
							{
								//console.log("3");
								return true;
							}
							else if(region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0)
							{
								//console.log("4");
								return true;
							}
						})).enter().append("g").attr("class", "arc");
		


		g.append("path").attr("d", function (d){
			//console.log(d.data.Y1991+": "+arc(d));
			return arc(d);})
			.style("fill", function(d) {
			if (region == 1 && $.inArray(d.data.Region, hamptonRoadRegions) >= 0) {
				//console.log(d.data.Region);
				return color(d.data.Region);
			}
			else if(region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0)
			{
				//console.log(d.data.Region);
				return color(d.data.Region);
			}
			else if(region == 3 && $.inArray(d.data.Region, dcRegions) >= 0)
			{
				//console.log("3");
				return color(d.data.Region);
			}
			else if(region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0)
			{
				//console.log("4");
				return color(d.data.Region);
			}
		});

		g.append("text").attr("transform", function(d) {
			return "translate(" + labelArc.centroid(d) + ")";
		}).attr("dy", ".10em").text(
				function(d) {
					var value=0;
					if (region == 1 && $.inArray(d.data.Region, hamptonRoadRegions) >= 0) {
						value = Math.round(10000 * (getPieGraphValue(x,d.data))) / 100;
					}
					else if(region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0)
					{
						value = Math.round(10000 * (getPieGraphValue(x,d.data))) / 100;
					}
					else if(region == 3 && $.inArray(d.data.Region, dcRegions) >= 0)
					{
						value = Math.round(10000 * (getPieGraphValue(x,d.data))) / 100;
					}
					else if(region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0)
					{
						value = Math.round(10000 * (getPieGraphValue(x,d.data))) / 100;
					}
					console.log("val---"+parseFloat(value).toFixed(1));
					return parseFloat(value).toFixed(1) + "%";
				});

		var legend = d3.select("#chartHotelPieLegend").append("svg").attr(
				"class", "legend").attr("width", radius * 2).attr("height",
				radius * 2).selectAll("g").data(data.filter(function(d) {
			if (region == 1 && $.inArray(d.Region, hamptonRoadRegions) >= 0) {
					return true;
			}
			else if(region == 2 && $.inArray(d.Region, virginiaRegions) >= 0)
			{
				return true;
			}
			else if(region == 3 && $.inArray(d.Region, dcRegions) >= 0)
			{
				
				return true;
			}
			else if(region == 4 && $.inArray(d.Region, dcRoomRegions) >= 0)
			{
				return true;
			}
		})).enter().append("g").attr("transform", function(data, i) {
			if (region == 1 && $.inArray(data.Region, hamptonRoadRegions) >= 0) {
				return "translate(0," + i * 20 + ")";
			}
			else if(region == 2 && $.inArray(data.Region, virginiaRegions) >= 0)
			{
				return "translate(0," + i * 20 + ")";
			}
			else if(region == 3 && $.inArray(data.Region, dcRegions) >= 0)
			{
				return "translate(0," + i * 20 + ")";
			}
			else if(region == 4 && $.inArray(data.Region, dcRoomRegions) >= 0)
			{
				return "translate(0," + i * 20 + ")";
			}
		});

		legend.append("rect").attr("width", 18).attr("height", 18).style(
				"fill", function(data, i) {
					if (region == 1 && $.inArray(data.Region, hamptonRoadRegions) >= 0) {
							return color(data.Region);
					}
					else if(region == 2 && $.inArray(data.Region, virginiaRegions) >= 0)
					{
						return color(data.Region);
					}
					else if(region == 3 && $.inArray(data.Region, dcRegions) >= 0)
					{
						return color(data.Region);
					}
					else if(region == 4 && $.inArray(data.Region, dcRoomRegions) >= 0)
					{
						return color(data.Region);
					}
				});

		legend.append("text").attr("x", 24).attr("y", 9).attr("dy", ".35em")
				.text(function(data) {
					if (region == 1 && $.inArray(data.Region, hamptonRoadRegions) >= 0) {
							return data.Region;
					}
					else if(region == 2 && $.inArray(data.Region, virginiaRegions) >= 0)
					{
						return data.Region;
					}
					else if(region == 3 && $.inArray(data.Region, dcRegions) >= 0)
					{
						return data.Region;
					}
					else if(region == 4 && $.inArray(data.Region, dcRoomRegions) >= 0)
					{
						return data.Region;
					}
				});
	});

	//current year pie chart
	var pie1 = d3.layout.pie().sort(null).value(function(d) {
		if (region == 1 && $.inArray(d.Region, hamptonRoadRegions) >= 0) {
				return d["Y"+(new Date().getFullYear()-1)];
		}
		else if (region == 2 && $.inArray(d.Region, virginiaRegions) >= 0) {
			return d["Y"+(new Date().getFullYear()-1)];
	}
		else if (region == 3 && $.inArray(d.Region, dcRegions) >= 0) {
		//	console.log("1-"+"Y"+new Date().getFullYear()-1);
				return d["Y"+(new Date().getFullYear()-1)];
		}
		else if (region == 4 && $.inArray(d.Region, dcRoomRegions) >= 0) {
			return d["Y"+(new Date().getFullYear()-1)];
	}
		else
			{
			return 0;
			}
			
	});

	var svg1 = d3.select("#chartHotelPieSecond").append("svg").attr("width",
			width).attr("height", height).append("g").attr("transform",
			"translate(" + width / 2 + "," + height / 2 + ")");
	
	svg1.append("text")
    .attr("x", 10)             
    .attr("y", -130)
    .attr("text-anchor", "middle")  
    .style("font-size", "12px") 
    .style("text-decoration", "underline")
    .style("font-weight", "bold")  
    .text(new Date().getFullYear()-1);


	d3.csv("HotelData_PieGraph.csv", function(error, data) {
		if (error)
			throw error;

		var g = svg1.selectAll(".arc").data(pie1(data).filter(function(d) {
			if (region == 1 && $.inArray(d.data.Region, hamptonRoadRegions) >= 0) {
					return true;
			}
			else if (region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0) {
					return true;
			}
			else if (region == 3 && $.inArray(d.data.Region, dcRegions) >= 0) {
				return true;
		}
			else if (region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0) {
				return true;
		}
		})).enter().append("g").attr("class", "arc");

		g.append("path").attr("d", arc).style("fill", function(d) {
			if (region == 1 && $.inArray(d.data.Region, hamptonRoadRegions) >= 0) {
					return color(d.data.Region);
			}
			else if (region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0) {
				return color(d.data.Region);
			}
			else if (region == 3 && $.inArray(d.data.Region, dcRegions) >= 0) {
				return color(d.data.Region);
			}
			else if (region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0) {
				return color(d.data.Region);
			}
		});

		g.append("text").attr("transform", function(d) {
			if (region == 1 && $.inArray(d.data.Region, hamptonRoadRegions) >= 0) {
				//console.log(labelArc.centroid(d));
				return "translate(" + labelArc.centroid(d) + ")";
			}
			else if (region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0) {
				return "translate(" + labelArc.centroid(d) + ")";
			}
			else if (region == 3 && $.inArray(d.data.Region, dcRegions) >= 0) {
				return "translate(" + labelArc.centroid(d) + ")";
			}
			else if (region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0) {
				return "translate(" + labelArc.centroid(d) + ")";
			}
		}).attr("dy", ".10em").text(function(d) {
			var value=0;
			if (region == 1 && $.inArray(d.data.Region, hamptonRoadRegions) >= 0) {
				value= Math.round(10000 * (d.data["Y"+(new Date().getFullYear()-1)])) / 100;
			}
			else if (region == 2 && $.inArray(d.data.Region, virginiaRegions) >= 0) {
				value= Math.round(10000 * (d.data["Y"+(new Date().getFullYear()-1)])) / 100;
			}
			else if (region == 3 && $.inArray(d.data.Region, dcRegions) >= 0) {
				//console.log("data--"+d.data["Y"+(new Date().getFullYear()-1)]);
				value= Math.round(10000 * (d.data["Y"+(new Date().getFullYear()-1)])) / 100;
			}
			else if (region == 4 && $.inArray(d.data.Region, dcRoomRegions) >= 0) {
				value= Math.round(10000 * (d.data["Y"+(new Date().getFullYear()-1)])) / 100;
			}
			return parseFloat(value).toFixed(1) + "%"
		});
	});
}

function drawHotelCharts(area, gtype) {
	$("#hotelRevCharts").empty();

	// console.log(area+":"+gtype);

	var margin = {
		top : 20,
		right : 20,
		bottom : 60,
		left : 60
	}, width = screen.width - margin.left - margin.right - 400, height = screen.height
			- margin.top - margin.bottom - 400;

	var x0 = d3.scale.ordinal().rangeRoundBands([ 0, width ], .1);

	var x1 = d3.scale.ordinal();

	var y = d3.scale.linear().range([ height, 0 ]);

	var tip;

	var color;

	if (gtype == "Revenue") {
		color = d3.scale.ordinal().range([ "#9966ff", "#99ff66" ]);
	} else if (gtype == "ADR") {
		color = d3.scale.ordinal().range([ "#9999ff", "#ff0066" ]);
	} else if (gtype == "Revpar") {
		color = d3.scale.ordinal().range([ "#f7786b", "#034f84" ]);
	} else if (gtype == "ORate") {
		color = d3.scale.ordinal().range([ "#cc6699" ]);
	} else {
		color = d3.scale.ordinal().range([ "#ffcc66", "#33ccff" ]);
	}

	var xAxis = d3.svg.axis().scale(x0).orient("bottom");

	var yAxis = d3.svg.axis().scale(y).orient("left").tickFormat(function(d) {
		if (gtype == "Revenue") {
			return "$"+parseFloat(d).toFixed(1);
		} else if (gtype == "ADR") {
			return "$"+parseFloat(d).toFixed(1);
		} else if (gtype == "Revpar") {
			return "$"+parseFloat(d).toFixed(1);
		} else if (gtype == "ORate") {
			return parseFloat(d).toFixed(1) + "%";
		} else {
			return parseFloat(d).toFixed(1).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");;
		}
		});

	var svg = d3.select("#hotelRevCharts").append("svg").attr("width",
			width + margin.left + margin.right).attr("height",
			height + margin.top + margin.bottom).append("g").attr("transform",
			"translate(" + margin.left + "," + margin.top + ")");
	
	var heading="";
	if (gtype == "Revenue") {
		if(area == "US" || area =="VA" || area =="DC")
			{
			heading="Billions of $";
			}
		else
			{
			heading="Millions of $";
			}
		
	} else if (gtype == "ADR") {
		heading="";
	} else if (gtype == "Revpar") {
		heading="";
	} else if (gtype == "ORate") {
		heading="";
	} else {
		heading="In Thousands";
	}
	
	svg.append("text")
    .attr("x", (width / 2))             
    .attr("y", 0 - (margin.top / 2))
    .attr("text-anchor", "middle")  
    .style("font-size", "12px") 
    .style("text-decoration", "underline")  
    .text(heading);
	
	

	d3.csv("HotelData_Graph.csv",function(error, data) {
		if (error)
			throw error;

		var ageNames = d3.keys(data[0])
			.filter(function(key) {
				if (key !== "Year") {
					if (gtype == "AO") {
						if (key == area+ "Available" || key == area + "Occupied") {
							return key;
							}
						} else if (gtype == "ORate")
						{
							if (key == area + "ORate") 
							{
								return key;
								}
							} else {
								if (key == area + "Real" + gtype || key == area + "" + gtype) {
														return key;
										}
									}

								}
						});

						data.forEach(function(d) {
							d.ages = ageNames.map(function(name) {
								return {name : name,value : +d[name]};
							});
						});

						x0.domain(data.map(function(d) {
							if(gtype=="Revenue" || gtype=="AO")
								{
								if(d.Year !="YTD2015" && d.Year !="YTD2016")
									{
									//console.log(d.value);
										return d.Year;
									}
								}
							else
								{
								return d.Year;
								}
						}));

						tip = d3
								.tip()
								.attr('class', 'd3-tip')
								.offset([ -10, 0 ])
								.html(
										function(d) {
											return value = "<strong>Value:</strong> <span style='color:white'>"
													+ d.value + "</span>";
										});

						x1.domain(ageNames).rangeRoundBands(
								[ 0, x0.rangeBand() ]);
						y.domain([ 0, d3.max(data, function(d) {
							return d3.max(d.ages, function(d) {
									return d.value;
							});
						}) ]);

						svg.append("g").attr("class", "x axis").attr(
								"transform", "translate(0," + height + ")")
								.call(xAxis).selectAll("text").attr("y", -1)
								.attr("x", -35).attr("dy", ".35em").attr(
										"transform", "rotate(270)").style(
										"text-anchor", "bottom");

						svg.append("g").attr("class", "y axis").call(yAxis)
								.append("text")
								.attr("transform", "rotate(-90)").attr("y", 6)
								.attr("dy", ".71em")
								.style("text-anchor", "end").text("");

						svg.call(tip);

						var state = svg.selectAll(".state").data(data).enter()
								.append("g").attr("class", "state").attr(
										"transform",
										function(d) {
											return "translate(" + x0(d.Year)
													+ ",0)";
										});

						state.selectAll("rect").data(function(d) {
							return d.ages;
						}).enter().append("rect").attr("width",20)
								.attr("x", function(d) {
									return x1(d.name);
								}).attr("y", function(d) {
									//console.log(d.name+":"+d.value);
									return y(d.value);
								}).attr("height", function(d) {
									return height - y(d.value);
								}).style("fill", function(d) {
									return color(d.name);
								}).on('mouseover', tip.show).on('mouseout',
										tip.hide);

						var legend = svg.selectAll(".legend").data(
								ageNames.slice().reverse()).enter().append("g")
								.attr("class", "legend").attr(
										"transform",
										function(d, i) {
											return "translate(0," + i * 20
													+ ")";
										});

						legend.append("rect").attr("x", 100).attr(
								"width", 18).attr("height", 18).style("fill",
								color);

						legend.append("text").attr("x", 95)
								.attr("y", 9).attr("dy", ".35em").style(
										"text-anchor", "end").text(function(d) {
									return d;
								});

					});
}