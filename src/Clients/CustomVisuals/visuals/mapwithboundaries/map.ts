declare var Microsoft;

var map, d3MapTools, countiesLayer, counties;

var svg;

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

    export class APDistrictMap implements IVisual {        
        public static capabilities: VisualCapabilities = {
            dataRoles: [
                {
                    name: 'Category',
                    kind: VisualDataRoleKind.Grouping,
                    displayName: data.createDisplayNameGetter('Role_DisplayName_Location'),
                    preferredTypes: [
                        { geography: { address: true } },
                        { geography: { city: true } },
                        { geography: { continent: true } },
                        { geography: { country: true } },
                        { geography: { county: true } },
                        { geography: { place: true } },
                        { geography: { postalCode: true } },
                        { geography: { region: true } },
                        { geography: { stateOrProvince: true } },
                    ],
                },
                {
                    name: 'Height',
                    kind: VisualDataRoleKind.Measure,
                    displayName: 'Bar Height',
                },
                {
                    name: 'Heat',
                    kind: VisualDataRoleKind.Measure,
                    displayName: 'Heat Intensity',
                }
            ],
            objects: {
                general: {
                    displayName: data.createDisplayNameGetter('Visual_General'),
                    properties: {
                        formatString: {
                            type: { formatting: { formatString: true } },
                        },
                    },
                },
                legend: {
                    displayName: data.createDisplayNameGetter('Visual_Legend'),
                    properties: {
                        show: {
                            displayName: data.createDisplayNameGetter('Visual_Show'),
                            type: { bool: true }
                        },
                        position: {
                            displayName: data.createDisplayNameGetter('Visual_LegendPosition'),
                            type: { formatting: { legendPosition: true } }
                        },
                        showTitle: {
                            displayName: data.createDisplayNameGetter('Visual_LegendShowTitle'),
                            type: { bool: true }
                        },
                        titleText: {
                            displayName: data.createDisplayNameGetter('Visual_LegendTitleText'),
                            type: { text: true }
                        }
                    }
                },
                dataPoint: {
                    displayName: data.createDisplayNameGetter('Visual_DataPoint'),
                    properties: {
                        defaultColor: {
                            displayName: data.createDisplayNameGetter('Visual_DefaultColor'),
                            type: { fill: { solid: { color: true } } }
                        },
                        showAllDataPoints: {
                            displayName: data.createDisplayNameGetter('Visual_DataPoint_Show_All'),
                            type: { bool: true }
                        },
                        fill: {
                            displayName: data.createDisplayNameGetter('Visual_Fill'),
                            type: { fill: { solid: { color: true } } }
                        },
                        fillRule: {
                            displayName: data.createDisplayNameGetter('Visual_Gradient'),
                            type: { fillRule: {} },
                            rule: {
                                inputRole: 'Gradient',
                                output: {
                                    property: 'fill',
                                    selector: ['Category'],
                                },
                            },
                        }
                    }
                },
                categoryLabels: {
                    displayName: data.createDisplayNameGetter('Visual_CategoryLabels'),
                    properties: {
                        show: {
                            displayName: data.createDisplayNameGetter('Visual_Show'),
                            type: { bool: true }
                        },
                        color: {
                            displayName: data.createDisplayNameGetter('Visual_LabelsFill'),
                            type: { fill: { solid: { color: true } } }
                        },
                    },
                },
            },
            dataViewMappings: [{
                conditions: [
                    { 'Category': { max: 1 }, 'Height': { max: 1 }, 'Heat': { max: 1 } },
                ],
                categorical: {
                    categories: {
                        for: { in: 'Category' },
                        dataReductionAlgorithm: { top: {} }
                    },
                    values: {
                        select: [
                            { bind: { to: 'Height' } },
                            { bind: { to: 'Heat' } },
                        ]
                    },
                    rowCount: { preferred: { min: 2 } }
                },
            }],
            sorting: {
                custom: {},
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
        //private svg: D3.Selection;
        
        /** This is called once when the visual is initialially created */
        public init(options: VisualInitOptions): void {
            this.colorPalette = options.style.colorPalette.dataColors;
            // element is the element in which your visual will be hosted.
            this.hostContainer = options.element.css('overflow-x', 'hidden');                                                                                                                                                                                                                 
        this.content = $("<div>").appendTo(options.element).css("position", "relative");
            
       this.loadScripts(() => {
                               map = new Microsoft.Maps.Map(this.content[0], {
                                                credentials: 'AsLj_okjJSBjbOnCP3C7E_opWa8qmtwmWV69nblODwur1a7Hq0_G4SWbm9rcpUgq',
                                                center: new Microsoft.Maps.Location(16.5000, 80.6400),
                                                zoom: 6
               });	
		Microsoft.Maps.Events.addHandler(map, 'tiledownloadcomplete', () => this.loadCounties());                           			   
        });
        }
        
        public loadCounties() {
            d3MapTools = new D3OverlayManager(map);
						if(countiesLayer != null)
                        {                            
                            if(countiesLayer.svg[0][0].parentElement != null)
                            {
                            countiesLayer.svg[0][0].parentElement.remove();                       }
                        }
            countiesLayer = d3MapTools.addLayer({
                loaded: function (svg, projection) {
					
					var jsondata = "";
                    var xhr = new XMLHttpRequest();
					xhr.open('GET', encodeURI('https://icrisat.blob.core.windows.net/json/apdists.json'), false);
					xhr.onload = function() {
					if (xhr.status === 200) {
					//alert('User\'s name is ' + xhr.responseText);
						jsondata = xhr.responseText;
						debugger;
						//d3.json(JSON.parse(jsondata), function (error, Topology) {
                        //var Topology = JSON.parse(jsondata);                          
                        //var topoData = topojson.feature(Topology, Topology.objects.counties).features;
						
						var formattedTopo = JSON.parse(jsondata);
						console.log(svg);
					 counties = svg.selectAll('path')
                            .data(formattedTopo)
                          .enter().append('path')
                            .attr('class', 'counties')
                            .attr('d', projection)
                        .on('click', function (feature) {
                            debugger;
                            alert(feature.properties["Mandal"]);
                        });	
						
						//});
					}
					else {
					alert('Request failed.  Returned status of ' + xhr.status);
					}
					};
					xhr.send();
					
                }
            });
        }
        
        public loadScripts(onScriptsLoaded: () => void) {		
            if(APDistrictMap.scriptsLoaded) {
                onScriptsLoaded();
                return;
            }
            
            APDistrictMap.scriptsLoaded = true;
            $.when(
                                $.getScript('https://d3js.org/d3.v3.min.js'),
                                $.getScript('https://cdnjs.cloudflare.com/ajax/libs/topojson/1.6.20/topojson.js'),
                                $.getScript("https://ecn.dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=7.0&s=1"),
                                $.Deferred(( deferred ) => $( deferred.resolve ))
                    ).done(() => {
                setTimeout(() => {
                    Microsoft.Maps.registerModule("D3OverlayModule", D3OverlayManager);
                    Microsoft.Maps.loadModule("D3OverlayModule"); 
					debugger;					
                    onScriptsLoaded();					
                }, 0);
            });
        }

        /** Update is called for data updates, resizes & formatting changes */
        public update(options: VisualUpdateOptions) {
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

            if (!dataViews) return;

            this.updateContainerViewports(options.viewport);

            var viewModel = APDistrictMap.converter(dataViews[0], this.colorPalette);
            var transposedSeries = d3.transpose(viewModel.values.map(d => d.values.map(d => d)));            
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

            _center = map.getCenter();
            _zoom = map.getZoom();			
			//var rect1 = map.getRootElement().getBoundingClientRect();			
            _mapWidth = map.getWidth();
            _mapHeight = map.getHeight();

            _d3Projection = d3.geo.path().projection(<any>_mapProjection);

            Microsoft.Maps.Events.addHandler(map, 'viewchange', function () {
                var z = map.getZoom();

                if (z != _zoom) {
                    _d3Container.style.display = 'none';
                } else {
                    _d3Container.style.display = '';

                    var cPixel = map.tryLocationToPixel(_center, Microsoft.Maps.PixelReference.viewport);
                    _d3Container.style.top = cPixel.y + 'px';
                    _d3Container.style.left = cPixel.x + 'px';
                }
            });

            Microsoft.Maps.Events.addHandler(map, 'viewchangeend', function () {
                _d3Container.style.display = '';

                _d3Container.style.top = '0px';
                _d3Container.style.left = '0px';

                _worldWidth = Math.pow(2, map.getZoom()) * 256;
                _prevX = { x: 0, y: 0 };
				
				
                if (_mapHeight != map.getHeight() || _mapWidth != map.getWidth()) {
                    //Map was resized. Resize layers.
                    _mapWidth = map.getWidth();
                    _mapHeight = map.getHeight();
                }
var rect = map.getRootElement().getBoundingClientRect();
                for (var i = 0; i < _layers.length; i++) {
                    if (_layers[i].svg) {
                        //Enusre layers are the same size as the map.
                        _layers[i].svg.attr("width", _mapWidth);
                        _layers[i].svg.attr("height", _mapHeight);
                        _layers[i].svg[0][0].style.left = - rect.width / 2 + 'px';
                        _layers[i].svg[0][0].style.top = - rect.height / 2 + 'px';

                        //Update data to align with projection
                        _layers[i].svg.selectAll('path').attr("d", _d3Projection);
                    }
                }

                _center = map.getCenter();
                _zoom = map.getZoom();
            });
        } else {
            //Map not fully loaded yet, try again in 100ms
            setTimeout(_init, 100);
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

    this.removeLayer = function (layer) {
		debugger;
        for (var i = 0; i < _layers.length; i++) {
            if (_layers[i].id == layer.id) {
                _d3Container.removeChild(layer.svg[0][0]);
                _layers.splice(i, 1);
                break;
            }
        }
    };

    _init();
};

// Call the Module Loaded method
//Microsoft.Maps.moduleLoaded('D3OverlayModule');

