declare var Microsoft;

module powerbi.visuals {    
    export interface CategoryViewModel {
        value: string;
        identity: string;
        color: string;
    }

    export interface ValueViewModel {
        values: any[];
    }

    export interface ViewModel {
        categories: CategoryViewModel[];
        values: ValueViewModel[];
    }

    export class RegionMap implements IVisual {        
        public static capabilities: VisualCapabilities = {
            dataRoles: [
                {
                    name: 'DISTRICT',
                    kind: VisualDataRoleKind.Grouping,
                    displayName: 'DISTRICT',                    
                },
                {
                    name: 'MANDAL',
                    kind: VisualDataRoleKind.Grouping,
                    displayName: 'MANDAL', 
                }
            ],            
            dataViewMappings: [{
                conditions: [
                    { 'DISTRICT': { max: 1 } },
                ],
                categorical: {
                    categories: {
                        for: { in: 'DISTRICT' },
                        dataReductionAlgorithm: { top: {} }
                    },
                    values: {
                        select: [
                            { bind: { to: 'MANDAL' } },
                        ]
                    },
                    rowCount: { preferred: { min: 1 } }
                },
            }],
            sorting: {
                custom: {},
            },
			drilldown: {
            roles: ['DISTRICT']
			}
        };    
		
		public static converter(dataView: DataView, colors: IDataColorPalette): ViewModel {
            var viewModel: ViewModel = {
                categories: [],
                values: []
            }
            if (dataView) {
                var categorical = dataView.categorical;
                if (categorical) {
                    var categories = categorical.categories;
                    var series = categorical.values;
                    var formatString = dataView.metadata.columns[0].format;

                    if (categories && series && categories.length > 0 && series.length > 0) {
                        for (var i = 0, catLength = categories[0].values.length; i < catLength; i++) {
                            viewModel.categories.push({
                                color: colors.getColorByIndex(i).value,
                                value: categories[0].values[i],
                                identity: ''
                            })

                            for (var k = 0, seriesLength = series.length; k < seriesLength; k++) {
                                var value = series[k].values[i];
                                if (k == 0) {
                                    viewModel.values.push({ values: [] });
                                }
                                viewModel.values[i].values.push(value);
                            }
                        }
                    }
                }
            }

            return viewModel;
        }
		
		private static scriptsLoaded: boolean = false;        		
        private hostContainer: JQuery;
		private colorPalette: IDataColorPalette;
		private content: JQuery;
		private geoJson: any;
        private d3MapTools: any;
        private map: any;
        private regionLayer: any;
        private selectedDistricts = '';
        private selectedMandals = '';
		private static isMapLoaded: boolean = false;		
		private mapState = 'UnInitiated';	
		private preSelectedDistricts = '';
        private preSelectedMandals = '';
		private defaultGeoJson: any;
        //private svg: D3.Selection;
		private mapSvg: any;
		private mapProjection: any;		
        
        /** This is called once when the visual is initialially created */
        public init(options: VisualInitOptions): void {
		if(this.mapState === 'UnInitiated')
		{			
            this.colorPalette = options.style.colorPalette.dataColors;
            // element is the element in which your visual will be hosted.
            this.hostContainer = options.element.css('overflow-x', 'hidden');                                                                                                                                                                                                                 
			this.content = $("<div>").appendTo(options.element).css("position", "relative");
			this.mapState = 'Initiated';
			this.loadScripts(() => {
               debugger;
			   this.map = new Microsoft.Maps.Map(this.content[0], {
                                                credentials: 'AsLj_okjJSBjbOnCP3C7E_opWa8qmtwmWV69nblODwur1a7Hq0_G4SWbm9rcpUgq',
                                                center: new Microsoft.Maps.Location(16.5000, 80.6400),
                                                mapTypeId: Microsoft.Maps.MapTypeId.birdseye,
                                                zoom: 6
                    });
                    
                    var tiledownloadcompleteId = Microsoft.Maps.Events.addHandler(this.map, 'tiledownloadcomplete', () => {
                        Microsoft.Maps.Events.removeHandler(tiledownloadcompleteId);
						this.getDefaultGeoJson();						
						this.selectedDistricts = '';
						this.selectedMandals = ''
                        RegionMap.isMapLoaded = true;												
                    });                                 			 
        	});
			}
        }
		
		//get Default GeoJson from Web API
		private getDefaultGeoJson()
		{			
				debugger;
				var id = "";				
				
				$.ajax({
				   url: 'https://geojsonapi.azurewebsites.net/api/mandal/' + "?" + id,
				   async: true,
                   dataType: 'jsonp',   
                   crossDomain: true,                                     
				}).done(x => {					
                    this.defaultGeoJson = x;
					this.mapState = 'GeoJsonLoaded';

					this.d3MapTools = new D3OverlayManager(this.map);
			
					this.regionLayer = this.d3MapTools.addLayer({
						loaded: (svg, projection) => {
					
						this.mapSvg = svg;														
					                    
						this.mapProjection = projection;					
						}
					});
				
				}).fail(x => {
					console.log('Request failed.  Returned status of ' + x);
				});							
		}
        
		private getRegionGeoJson()
		{		
			debugger;
			if(this.defaultGeoJson !== null && this.defaultGeoJson !== undefined && this.defaultGeoJson !== '')
			{
			var filteredData:any[] = [];				
			
			var ListOfDistrcits = this.selectedDistricts; //"Nellore,Anantapur";
			var ListOfMandals = this.selectedMandals; //"Podalakur,Raptadu";	

			//var ListOfDistrcits = "Nellore,Anantapur";
			//var ListOfMandals = "Podalakur,Raptadu";
			
			for (var arrIndex in this.defaultGeoJson) 
            {
				if(ListOfDistrcits !== null && ListOfDistrcits !== undefined && ListOfDistrcits != '')
				{
					var lstDistricts = ListOfDistrcits.split(",");
					
					for (var distIndex in lstDistricts) 
					{
						if(ListOfMandals !== null && ListOfMandals !== undefined && ListOfMandals != '')
						{
							var lstMandals = ListOfMandals.split(",");
					
							for (var mandalIndex in lstMandals) 
							{
								if (this.defaultGeoJson[arrIndex].properties.District == lstDistricts[distIndex] && this.defaultGeoJson[arrIndex].properties.Mandal == lstMandals[mandalIndex])
								{									
									filteredData.push(this.defaultGeoJson[arrIndex]);
								}
							}
						}
						else
						{
							if (this.defaultGeoJson[arrIndex].properties.District == lstDistricts[distIndex])
							{
								filteredData.push(this.defaultGeoJson[arrIndex]);
							}
						}
						
					}
				}
				else
				{
					if(ListOfMandals !== null && ListOfMandals !== undefined && ListOfMandals != '')
						{
							var lstMandals = ListOfMandals.split(",");
					
							for (var mandalIndex in lstMandals) 
							{
								if (this.defaultGeoJson[arrIndex].properties.Mandal == lstMandals[mandalIndex])
								{
									filteredData.push(this.defaultGeoJson[arrIndex]);
								}
							}
						}
					else
						{
							filteredData.push(this.defaultGeoJson[arrIndex]);
						}
				}								
			}
			
            debugger;
			
            //console.log(filteredData);
                
			this.geoJson = filteredData;
			
			this.loadSelectedRegions();
			}
		}
		
        private loadSelectedRegions() {            												
				//debugger;												
				var div = d3.select("body").append("div")	
				.attr("class", "tooltip")				
				.style("opacity", 0);
				
				var updateSelection = this.mapSvg.selectAll("path").data(this.geoJson)				
				updateSelection.enter().append("path").attr('class', 'region')				
				updateSelection.attr('d', this.mapProjection)
				updateSelection.exit().remove()	

				updateSelection.on("mouseover", function(d) {				
										
				div.transition()		
                .duration(200)		
                .style("opacity", .9);		
				div.html(d.properties.Mandal)
				.style("left", (d3.event.clientX) + "px")		
                .style("top", (d3.event.clientY - 28) + "px");;	           

				}).on("mouseout", function() {
					 div.transition()		
                	.duration(500)		
                	.style("opacity", 0);
				})				
        }
		
        private loadScripts(onScriptsLoaded: () => void) {	
			//if(this.mapState === 'Initiated')
			//{
				if(RegionMap.scriptsLoaded) 
				{
					onScriptsLoaded();
					return;
				}
			            
            $.when(                                       
                    $.getScript("https://ecn.dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=7.0&s=1"),
                    $.Deferred(( deferred ) => $( deferred.resolve ))
                    ).done(() => {
						(function onModulesLoaded() {
							if(!Microsoft.Maps.registerModule) {                                
								return setTimeout(onModulesLoaded, 10);
							}
						
							Microsoft.Maps.registerModule("D3OverlayModule", D3OverlayManager);
							Microsoft.Maps.loadModule("D3OverlayModule");
							RegionMap.scriptsLoaded = true;                            
							onScriptsLoaded();
							})();					
					});
					this.mapState = 'ScriptsLoaded';
			//}
        }

        /** Update is called for data updates, resizes & formatting changes */
        public update(options: VisualUpdateOptions) {
			debugger;		
            var dataViews = options.dataViews;
             //debugger;
            /*
            this.svg.attr({
                'width': options.viewport.width,
                'height': options.viewport.height
            });
            */
			this.content.css({
                'width': options.viewport.width,
                'height': options.viewport.height
            });
			console.log(JSON.stringify(dataViews));
            if (!dataViews) return;

            this.updateContainerViewports(options.viewport);

            var viewModel = RegionMap.converter(dataViews[0], this.colorPalette);
            
            debugger;                      
			//if(this.mapState === 'ScriptsLoaded' || this.mapState === 'SelectedRegionsUpdated' || this.mapState === 'Updated')
			//{
			if(RegionMap.isMapLoaded) {
			if(options.dataViews[0] !== undefined)
			{
			//if(options.dataViews[0].categorical.categories[0].source.displayName === 'DISTRICT')
			 //{
            debugger;    
			this.selectedDistricts = '';			
            for (var distRow in options.dataViews[0].table.rows) 
            { 
                if(this.selectedDistricts === '' || this.selectedDistricts === undefined || this.selectedDistricts === null)
                {
                    this.selectedDistricts = options.dataViews[0].table.rows[distRow][0];
                }
                else
                {
					if(this.selectedDistricts.toString().indexOf(options.dataViews[0].table.rows[distRow][0]) === -1)
					{
						this.selectedDistricts = this.selectedDistricts + ',' + options.dataViews[0].table.rows[distRow][0];
					}
                }
            }			 
			//}
			
			//if(options.dataViews[0].categorical.categories[0].source.displayName === 'MANDAL')
			 //{
            debugger;     
			this.selectedMandals = '';			
            for (var distRow in options.dataViews[0].table.rows) 
            { 
                if(this.selectedMandals === '' || this.selectedMandals === undefined || this.selectedMandals === null)
                {
                    this.selectedMandals = options.dataViews[0].table.rows[distRow][1];
                }
                else
                {
                    this.selectedMandals = this.selectedMandals + ',' + options.dataViews[0].table.rows[distRow][1];
                }
            }			 
			//}
			//this.mapState = 'Updated';
			}			
			}            			
            //}
			var transposedSeries = d3.transpose(viewModel.values.map(d => d.values.map(d => d)));
			//if(this.mapState === 'Updated')
			//{
				//if(this.preSelectedDistricts != this.selectedDistricts || this.preSelectedMandals != this.selectedMandals)
				//{
					//this.preSelectedDistricts = this.selectedDistricts;
					//this.preSelectedMandals = this.selectedMandals;
					debugger;
					if(this.mapSvg !== null && this.mapSvg !== undefined && this.mapSvg !== '')
						{
							this.getRegionGeoJson();
						}
				//}
			//}
        }

        private updateContainerViewports(viewport: IViewport) {
            //var width = viewport.width;
            //var height = viewport.height;

            //this.tHead.classed('dynamic', width > 400);
            //this.tBody.classed('dynamic', width > 400);

            //this.hostContainer.css({
                //'height': height,
                //'width': width
            //});
            //this.table.attr('width', width);
        }
        
        private format(d: number){
            var prefix = d3.formatPrefix(d);
            return d3.round(prefix.scale(d),2) + ' ' +prefix.symbol
        }
    }
}

var D3OverlayManager = function (map) {

    var _d3Projection = null,
        _center,
        _zoom,
        _mapWidth,
        _mapHeight,
        _mapDiv,
        _d3Container,
        _layers = [],
        _idCnt = 0,
        _worldWidth = 512,
        _prevX = null;

    function _init() {
        _mapDiv = map.getRootElement();
        if (_mapDiv.childNodes.length >= 3 && _mapDiv.childNodes[2].childNodes.length >= 2) {
            _d3Container = document.createElement('div');
            _d3Container.style.position = 'absolute';
            _d3Container.style.top = '0px';
            _d3Container.style.left = '0px';
            _mapDiv.childNodes[2].childNodes[1].appendChild(_d3Container);

            //_center = map.getCenter();
            //_zoom = map.getZoom();			
			//var rect1 = map.getRootElement().getBoundingClientRect();			
            _mapWidth = map.getWidth();
            _mapHeight = map.getHeight();

            _d3Projection = d3.geo.path().projection(<any>_mapProjection);

            Microsoft.Maps.Events.addHandler(map, 'viewchange', viewChange);
            Microsoft.Maps.Events.addHandler(map, 'viewchangeend', viewChangeEnd);
            
        } else {
            //Map not fully loaded yet, try again in 100ms
            setTimeout(_init, 100);
        }
        
         function viewChange() {
            var z = map.getZoom();
            if (z != _zoom) {
                _d3Container.style.display = 'none';
            } else {
                _d3Container.style.display = '';
                var cPixel = map.tryLocationToPixel(_center, Microsoft.Maps.PixelReference.viewport);
                _d3Container.style.top = cPixel.y + 'px';
                _d3Container.style.left = cPixel.x + 'px';
            }
        }
        
        function viewChangeEnd() {
            _d3Container.style.display = '';

            _d3Container.style.top = '0px';
            _d3Container.style.left = '0px';

            _worldWidth = Math.pow(2, map.getZoom()) * 256;
            _prevX = { x: 0, y: 0 };
            
            _mapWidth = map.getWidth();
            _mapHeight = map.getHeight();
            
            var rect = _mapDiv.parentNode.getBoundingClientRect();
            for (var i = 0; i < _layers.length; i++) {
                if (_layers[i].svg) {
                    //Enusre layers are the same size as the map.
                    _layers[i].svg.attr("width", _mapWidth);
                    _layers[i].svg.attr("height", _mapHeight);
                    _layers[i].svg[0][0].style.left = -(rect.width / 2) + 'px';
                    _layers[i].svg[0][0].style.top = -(rect.height / 2) + 'px';

                    //Update data to align with projection
                    _layers[i].svg.selectAll('path').attr("d", _d3Projection);
                }
            }
            
            _center = map.getCenter();
            _zoom = map.getZoom();
        }
    }

    function _mapProjection(coordinates: D3.Geo.Projection): any {
        var loc = new Microsoft.Maps.Location(coordinates[1], coordinates[0]);
        var pixel = map.tryLocationToPixel(loc, Microsoft.Maps.PixelReference.control);

        //Workaround to ensure polygons are rendered properly when more than one world is visible.
        if (_prevX != null) {
            // Check for a gigantic jump to prevent wrapping around the world
            var dist = pixel.x - _prevX;
            if (Math.abs(dist) > (_worldWidth * 0.9)) {
                if (dist > 0) { pixel.x -= _worldWidth } else { pixel.x += _worldWidth }
            }
        }
        _prevX = pixel.x;

        return  [pixel.x, pixel.y];
    }

    /*******************
    * Public functions
    ********************/

    this.addLayer = function (options) {
         //debugger;
        //options = {
            //loaded: function (svg, projection) { },    //Callback that is fired when the layer has been loaded.
            //viewChanged: function (svg, projection) { }    //Callback that is fired when the view needs updating
        //}

        var _id = _idCnt++,
            layerInfo = {
                id: _idCnt,
                options: options,
                svg: null
            };                       

        function _initLayer() {
            if (_d3Container) {                                           
                layerInfo.svg = d3.select(_d3Container).append('svg')
                            .attr('width', _mapWidth)
                            .attr('height', _mapHeight);

                layerInfo.svg[0][0].style.position = 'relative';
                layerInfo.svg[0][0].style.left = -_mapWidth / 2 + 'px';
                layerInfo.svg[0][0].style.top = -_mapHeight / 2 + 'px';

                if (options && options.loaded) {
                    options.loaded(layerInfo.svg, _d3Projection);
                }
            } else {
                //Map not fully loaded yet, try again in 100ms
                setTimeout(_initLayer, 100);
            }
        }


        _initLayer();

        _layers.push(layerInfo);
         //debugger;         
        return layerInfo;
    };

    this.clearLayers = function () {
        _layers.splice(0, _layers.map(x => x.svg[0][0].parentElement.remove()).length);
    };

    _init();
};

// Call the Module Loaded method
//Microsoft.Maps.moduleLoaded('D3OverlayModule');
