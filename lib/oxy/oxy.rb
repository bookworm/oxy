require 'rubygems'
require 'rbosa'
require 'active_support'  

@ps = OSA.app('Adobe Photoshop CS5') 
@ps.settings.ruler_units = OSA::AdobePhotoshopCS5::E440::PIXEL_UNITS 
   
@ps.instance_eval do

## Convenience Methods
  def doc()
    return current_document
  end
  
## Add and Create Methods
  def create_document(options = {})
    make(OSA::AdobePhotoshopCS5::Document, nil, nil, {
    }.merge(options))
  end  
  
  # Group can either be a full group Hash with {:path => ['parent_group', 'child_group'] } or a just the group name.
  # If there is just a group name the method will do the effort for you and find the path to that group.
  def add_layer(name, kind, group = nil)
    kinds = %w(NORMAL GRADIENTFILL PATTERNFILL TEXT SOLIDFILL)  
    if group == nil
      do_javascript %(  
        var doc = app.activeDocument;
        var layer = doc.artLayers.add();
        layer.name = "#{name || ''}";
        layer.kind = LayerKind.#{kinds.detect {|k| k.downcase == kind} || 'NORMAL'};   
      )
      return current_document.art_layers[0]  
    end    
    
    if group.is_a?(String)
      group = find_layer_set(group)   
    elsif group.is_a?(Hash)
      group = group
    end  
     
    if !group[:depth].empty?
      group[:depth] -= 1 if group[:depth] > 0  
    else  
      group[:depth] = group.path.count - 1
    end         
    
    add_layer_js = <<-eos  
      #include '/Developer/xtools/xlib/stdlib.js';
      #include '/Developer/oxy/lib.js';        
     
      var docRef = app.activeDocument;
      var doc   = app.activeDocument;   
      var layer_sets = doc.layerSets;     
        
      var layer = doc.artLayers.add();
      layer.name = "#{name || ''}";
      layer.kind = LayerKind.#{kinds.detect {|k| k.downcase == kind} || 'NORMAL'};   
      
      // Group  
      var group;       
      var group_path = "#{group[:path].join(',')}";
      group_path = group_path.split(',');   
      var group_depth = #{group[:depth]};  
      
      group_loop_layer_sets = function(layer_set, depth)
      {  
        if(depth == undefined) {
          depth = 1;
        }   
        if(depth == (group_path.length - 1)) {    
          group = layer_set.layerSets.getByName(group_path[depth]);
        } 
        else
        {
          layer_set = layer_set.layerSets.getByName(group_path[depth]);
          group_loop_layer_sets(layer_set, depth++);
        }
      }    

      group_loop_layer_sets(layer_sets.getByName(group_path[0]));
      
      // create background layer -stdlib.js bug fix
      // var layerRefBackground = docRef.artLayers.add();
      // docRef.activeLayer.name = "BG";
      // docRef.activeLayer.isBackgroundLayer = true;         

      // move layer set "Pixel"
      moveLayer(layer, group);

      // remove Background layer
      // layerRefBackground.remove();
    eos
    retvalue = do_javascript(add_layer_js)
    return retvalue
  end
  
  # Group can either be a full group Hash with {:path => ['parent_group', 'child_group'] } or a just the group name.
  # If there is just a group name the method will do the effort for you and find the path to that group.
  def add_group(name, group = nil)
    if group == nil 
      do_javascript %(  
        var doc = app.activeDocument;
        var group = doc.layerSets.add(); 
        group.name = "#{name}";       
      ) 
      return current_document.layer_sets[0] 
    end 
    
    if group.is_a?(String)
      group = find_layer_set(group)   
      puts group.inspect
    elsif group.is_a?(Hash)
      group = group
    end  
     
    if group.has_key?(:depth)
      group[:depth] -= 1 if group[:depth] > 0  
    else  
      group[:depth] = group.path.count - 1
    end
    
    add_group_js = <<-eos  
      #include '/Developer/xtools/xlib/stdlib.js';
      #include '/Developer/oxy/lib.js'; 
     
      var docRef = app.activeDocument;
      var doc   = app.activeDocument;   
      var layer_sets = doc.layerSets; 
      
      var group = doc.layerSets.add(); 
      group.name = "#{name}";    
       
      // dest  
      var dest;       
      var dest_path = "#{group[:path].join(',')}";
      dest_path = dest_path.split(',');   
      var dest_depth = #{group[:depth]};
    
      dest_loop_layer_sets = function(layer_set, depth)
      {  
        if(depth == undefined) {
          depth = 1;
        }   
        if(depth == (dest_path.length - 1)) {    
          dest = layer_set.layerSets.getByName(dest_path[depth]);
        } 
        else
        {
          layer_set = layer_set.layerSets.getByName(dest_path[depth]);
          dest_loop_layer_sets(layer_set, depth++);
        }
      }   
        
      if(dest_depth == 0) {
        dest = layer_sets.getByName(dest_path[0]);
      } 
      else {
        dest_loop_layer_sets(layer_sets.getByName(dest_path[0]));   
      }

      // create background layer -stdlib.js bug fix
      // var layerRefBackground = docRef.artLayers.add();
      // docRef.activeLayer.name = "BG";
      // docRef.activeLayer.isBackgroundLayer = true;         

      // move layer set "Pixel"
      moveLayer(group, dest);

      // remove Background layer
      // layerRefBackground.remove();  
    eos
    retvalue = do_javascript(add_group_js)
    return retvalue
  end

## Find Methods
# Its important to relize that these return hashes with the path and depth + the layer obj. e.g
# {:path => '', :depth => , :obj => }. 
# Its assumed that if you want just the layer object you'll know where its at and will thus use layers array directly.
# i.e @ps.doc.art_layers.get_by_bame().

  @found = {:depth => 0, :path => [], :index_path => [] } 
  def find_layer_set(named, layer_sets = nil)
    retvalue = {}
    if @found[:depth] == 0 
      @found[:path] = [] 
      @found[:index_path] = []
    end 
    layer_sets = doc.layer_sets.to_a if layer_sets == nil 
    layer_sets.each_with_index do |layer_set, index|
      sub = layer_set.layer_sets.to_a  
      name = layer_set.name   
      if name == named  
        @found[:depth] = @found[:depth] + 1
        @found[:path].push name  
        @found[:index_path].push index
        retvalue[:obj] = layer_set  
        retvalue = retvalue.merge!(@found)
        @found = {:depth => 0, :path => [], :index_path => [] } 
        return retvalue
      elsif sub.count > 0     
        @found[:depth] = @found[:depth] + 1
        @found[:path].push name 
        @found[:index_path].push index
        return find_layer_set(named, sub)
      end
    end   
    @found = {:depth => 0, :path => [], :index_path => [] }
    return retvalue
  end  

  def find_layer(named = nil, layers = doc.art_layers.to_a, layer_sets = doc.layer_sets.to_a)
    retvalue = {}
    if layers != nil
      layers.each do |layer|   
        if layer.name == named
          retvalue[:obj]  = layer  
          retvalue = retvalue.merge(@found)  
          @found = {:depth => 0, :path => [], :index_path => [] } 
          return retvalue 
        end
      end 
    end  
    if layer_sets != nil
      if @found[:depth] == 0 
        @found[:path] = [] 
        @found[:index_path] = []
      end
      layer_sets.each_with_index do |layer_set,index|  
        layers = layer_set.art_layers.to_a
        name = layer_set.name 
        @found[:depth] = @found[:depth] + 1
        @found[:path].push name       
        @found[:index_path].push index
        if layers.count > 0 
          retvalue = find_layer(named, layers, nil)
          return retvalue if retvalue.has_key?(:obj)   
        end 
        sub = layer_set.layer_sets.to_a
        if sub.count > 0     
          retvalue = find_layer(named, nil, sub)  
          return retvalue if retvalue.has_key?(:obj)
        end
      end 
    end  
    @found = {:depth => 0, :path => [], :index_path => [] }
    return retvalue  
  end    

## Move Methods           

  # For documentation see move_layer_to_group()
  def move_layer_to_group(layer_name, group)
    doc  = @app.current_document    
    dest = find_layer_set(doc.layer_sets.to_a, group)   
    src = find_layer(doc.art_layers.to_a, doc.layer_sets.to_a, layer_name)

    src[:depth] -= 1 if src[:depth] > 0
    dest[:depth] -= 1 if dest[:depth] > 0
  
    move_js = <<-eos   
      #include '/Developer/xtools/xlib/stdlib.js';
      #include '/Developer/oxy/lib.js';
 
      var docRef = app.activeDocument;
      var doc   = app.activeDocument;   
      var layer_sets = doc.layerSets;
    
      // src    
      var src;       
      var src_path = "#{src[:path].join(',')}";
      src_path = src_path.split(',');   
      var src_depth = #{src[:depth]};
    
      // dest  
      var dest;       
      var dest_path = "#{dest[:path].join(',')}";
      dest_path = dest_path.split(',');   
      var dest_depth = #{dest[:depth]};
    
      dest_loop_layer_sets = function(layer_set, depth)
      {  
        if(depth == undefined) {
          depth = 1;
        }   
        if(depth == (dest_path.length - 1)) {    
          dest = layer_set.layerSets.getByName(dest_path[depth]);
        } 
        else
        {
          layer_set = layer_set.layerSets.getByName(dest_path[depth]);
          dest_loop_layer_sets(layer_set, depth++);
        }
      }    

      dest_loop_layer_sets(layer_sets.getByName(dest_path[0]));    
    
      if(src_depth == 0 ) {
        src = doc.artLayers.getByName("#{layer_name}");   
      } else {
        src_loop_layer_sets(layer_sets);
      }

      src_loop_layers = function(layer_sets, depth)
      {       
        if(depth == undefined) {
          depth = 0;
        }
        layer_set = layer_sets.getByName(src_path[depth]);  
       
        if((depth + 1) == src_path.length) {    
          src = layer_set.artLayers.getByName("#{layer_name}");
        }   
        else
        {  
          layer_sets = layer_set.layerSets;
          src_loop_layers(layer_sets, depth++);     
        }    
      }    

      // create background layer -stdlib.js bug fix
      // var layerRefBackground = docRef.artLayers.add();
      // docRef.activeLayer.name = "BG";
      // docRef.activeLayer.isBackgroundLayer = true;         

      // move layer set "Pixel"
      moveLayer(src, dest);
    eos
    retvalue = do_javascript(move_js)
    return retvalue
  end
  
  def move_group_to_group(src, dest)
    doc  = @app.current_document    
    src  = find_layer_set(doc.layer_sets.to_a, src)
    dest = find_layer_set(doc.layer_sets.to_a, dest)  
  
    src[:depth] -= 1 if src[:depth] > 0
    dest[:depth] -= 1 if dest[:depth] > 0
   
    move_js = <<-eos    
      #include '/Developer/xtools/xlib/stdlib.js';
      #include '/Developer/oxy/lib.js';
   
      var docRef = app.activeDocument;
      var doc   = app.activeDocument;   
      var layer_sets = doc.layerSets;     
    
      // src    
      var src;       
      var src_path = "#{src[:path].join(',')}";
      src_path = src_path.split(',');   
      var src_depth = #{src[:depth]};  
    
      // dest  
      var dest;       
      var dest_path = "#{dest[:path].join(',')}";
      dest_path = dest_path.split(',');   
      var dest_depth = #{dest[:depth]};    
    
      dest_loop_layer_sets = function(layer_sets)
      {  
        if(dest_depth == (dest_path.length - 1) ) {
          dest = layer_sets.getByName(dest_path[dest_depth]);
        } 
        else
        {
          for(var i = 0; i <= layer_sets.length; i++) 
          {  
            if(layer_sets[i].name == dest_path[dest_depth])
            { 
              layet_sets = layer_sets[i].layerSets;  
              dest_depth = dest_depth + 1;
              dest_loop_layer_sets(layer_sets);   
            }      
          } 
        }
      }  
    
      src_loop_layer_sets = function(layer_sets)
      { 
        if(src_depth == (src_path.length - 1)) {
          src = layer_sets.getByName(src_path[src_depth]);
        }  
        else
        {
          for(var i = 0; i <= layer_sets.length; i++) 
          {  
            if(layer_sets[i].name == src_path[src_depth])
            { 
              layet_sets = layer_sets[i].layerSets;  
              src_depth = src_depth + 1;
              src_loop_layer_sets(layer_sets);   
            }       
          }
        }
      }
    
      dest_loop_layer_sets(layer_sets);    
      src_loop_layer_sets(layer_sets);
      
      // create background layer -stdlib.js bug fix
      // var layerRefBackground = docRef.artLayers.add();
      // docRef.activeLayer.name = "BG";
      // docRef.activeLayer.isBackgroundLayer = true;         

      // move layer set "Pixel"
      moveLayer(src, dest);

      // remove Background layer
      // layerRefBackground.remove();      
  	eos
  	retvalue = do_javascript(move_js)
    return retvalue
  end  
  
## Draw Methods
# For any of these methods your going to need make sure you set the active layer before calling them.    

  def draw_rec(offset, width, height, color = nil)
    x  = offset['left']
    y  = offset['top'] 
    ry = y + height   
    rx = x + width
  
    str = <<-eos  
      var idMk = charIDToTypeID( "Mk  " );
          var desc50 = new ActionDescriptor();
          var idnull = charIDToTypeID( "null" );
              var ref27 = new ActionReference();
              var idcontentLayer = stringIDToTypeID( "contentLayer" );
              ref27.putClass( idcontentLayer );
          desc50.putReference( idnull, ref27 );
          var idUsng = charIDToTypeID( "Usng" );
              var desc51 = new ActionDescriptor();
              var idType = charIDToTypeID( "Type" );
              var idsolidColorLayer = stringIDToTypeID( "solidColorLayer" );
              desc51.putClass( idType, idsolidColorLayer );
              var idShp = charIDToTypeID( "Shp " );
                  var desc52 = new ActionDescriptor();
                  var idTop = charIDToTypeID( "Top " );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc52.putUnitDouble( idTop, idPxl, #{y} ); // Y
                  var idLeft = charIDToTypeID( "Left" );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc52.putUnitDouble( idLeft, idPxl, #{x} );  // X
                  var idBtom = charIDToTypeID( "Btom" );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc52.putUnitDouble( idBtom, idPxl, #{ry} ); // Y
                  var idRght = charIDToTypeID( "Rght" );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc52.putUnitDouble( idRght, idPxl, #{rx} ); // X
              var idRctn = charIDToTypeID( "Rctn" );
              desc51.putObject( idShp, idRctn, desc52 );
              var idStyl = charIDToTypeID( "Styl" );
                  var ref28 = new ActionReference();
                  var idStyl = charIDToTypeID( "Styl" );
                  ref28.putName( idStyl, "Shape 1 style" );
              desc51.putReference( idStyl, ref28 );
          var idcontentLayer = stringIDToTypeID( "contentLayer" );
          desc50.putObject( idUsng, idcontentLayer, desc51 );
      executeAction( idMk, desc50, DialogModes.NO );     
    eos
    do_javascript(str)       
   
    if color != nil 
      color_js = <<-eos
        var idsetd = charIDToTypeID( "setd" );
            var desc23 = new ActionDescriptor();
            var idnull = charIDToTypeID( "null" );
                var ref15 = new ActionReference();
                var idcontentLayer = stringIDToTypeID( "contentLayer" );
                var idOrdn = charIDToTypeID( "Ordn" );
                var idTrgt = charIDToTypeID( "Trgt" );
                ref15.putEnumerated( idcontentLayer, idOrdn, idTrgt );
            desc23.putReference( idnull, ref15 );
            var idT = charIDToTypeID( "T   " );
                var desc24 = new ActionDescriptor();
                var idClr = charIDToTypeID( "Clr " );
                    var desc25 = new ActionDescriptor();
                    var idRd = charIDToTypeID( "Rd  " );
                    desc25.putDouble( idRd, #{color[:red]} );
                    var idGrn = charIDToTypeID( "Grn " );
                    desc25.putDouble( idGrn, #{color[:green]} );
                    var idBl = charIDToTypeID( "Bl  " );
                    desc25.putDouble( idBl, #{color[:blue]} );
                var idRGBC = charIDToTypeID( "RGBC" );
                desc24.putObject( idClr, idRGBC, desc25 );
            var idsolidColorLayer = stringIDToTypeID( "solidColorLayer" );
            desc23.putObject( idT, idsolidColorLayer, desc24 );
        executeAction( idsetd, desc23, DialogModes.NO );    
      eos
      do_javascript(color_js)       
    end      
  
    if name != nil    
      name_js = <<-eos
        var idsetd = charIDToTypeID( "setd" );
            var desc7 = new ActionDescriptor();
            var idnull = charIDToTypeID( "null" );
                var ref3 = new ActionReference();
                var idLyr = charIDToTypeID( "Lyr " );
                var idOrdn = charIDToTypeID( "Ordn" );
                var idTrgt = charIDToTypeID( "Trgt" );
                ref3.putEnumerated( idLyr, idOrdn, idTrgt );
            desc7.putReference( idnull, ref3 );
            var idT = charIDToTypeID( "T   " );
                var desc8 = new ActionDescriptor();
                var idNm = charIDToTypeID( "Nm  " );
                desc8.putString( idNm, "#{name}");
            var idLyr = charIDToTypeID( "Lyr " );
            desc7.putObject( idT, idLyr, desc8 );
        executeAction( idsetd, desc7, DialogModes.NO );
      eos
      do_javascript(name_js)
    end
  end      

  def draw_rounded_rec(offset, width, height, radius, color = nil)
    x  = offset['left']
    y  = offset['top'] 
    ry = y + height   
    rx = x + width        
  
    str = <<-eos
      var idMk = charIDToTypeID( "Mk  " );
          var desc77 = new ActionDescriptor();
          var idnull = charIDToTypeID( "null" );
              var ref42 = new ActionReference();
              var idcontentLayer = stringIDToTypeID( "contentLayer" );
              ref42.putClass( idcontentLayer );
          desc77.putReference( idnull, ref42 );
          var idUsng = charIDToTypeID( "Usng" );
              var desc78 = new ActionDescriptor();
              var idType = charIDToTypeID( "Type" );
              var idsolidColorLayer = stringIDToTypeID( "solidColorLayer" );
              desc78.putClass( idType, idsolidColorLayer );
              var idShp = charIDToTypeID( "Shp " );
                  var desc79 = new ActionDescriptor();
                  var idTop = charIDToTypeID( "Top " );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc79.putUnitDouble( idTop, idPxl, #{y} );
                  var idLeft = charIDToTypeID( "Left" );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc79.putUnitDouble( idLeft, idPxl, #{x} );
                  var idBtom = charIDToTypeID( "Btom" );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc79.putUnitDouble( idBtom, idPxl, #{ry} );
                  var idRght = charIDToTypeID( "Rght" );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc79.putUnitDouble( idRght, idPxl, #{rx} );
                  var idRds = charIDToTypeID( "Rds " );
                  var idPxl = charIDToTypeID( "#Pxl" );
                  desc79.putUnitDouble( idRds, idPxl, #{radius.gsub('px', '')} );  // Border radius
              var idRctn = charIDToTypeID( "Rctn" );
              desc78.putObject( idShp, idRctn, desc79 );
          var idcontentLayer = stringIDToTypeID( "contentLayer" );
          desc77.putObject( idUsng, idcontentLayer, desc78 );
      executeAction( idMk, desc77, DialogModes.NO ); 
    eos
    do_javascript(str)     
  
    if color != nil 
      color_js = <<-eos
        var idsetd = charIDToTypeID( "setd" );
            var desc23 = new ActionDescriptor();
            var idnull = charIDToTypeID( "null" );
                var ref15 = new ActionReference();
                var idcontentLayer = stringIDToTypeID( "contentLayer" );
                var idOrdn = charIDToTypeID( "Ordn" );
                var idTrgt = charIDToTypeID( "Trgt" );
                ref15.putEnumerated( idcontentLayer, idOrdn, idTrgt );
            desc23.putReference( idnull, ref15 );
            var idT = charIDToTypeID( "T   " );
                var desc24 = new ActionDescriptor();
                var idClr = charIDToTypeID( "Clr " );
                    var desc25 = new ActionDescriptor();
                    var idRd = charIDToTypeID( "Rd  " );
                    desc25.putDouble( idRd, #{color[:red]} );
                    var idGrn = charIDToTypeID( "Grn " );
                    desc25.putDouble( idGrn, #{color[:green]} );
                    var idBl = charIDToTypeID( "Bl  " );
                    desc25.putDouble( idBl, #{color[:blue]} );
                var idRGBC = charIDToTypeID( "RGBC" );
                desc24.putObject( idClr, idRGBC, desc25 );
            var idsolidColorLayer = stringIDToTypeID( "solidColorLayer" );
            desc23.putObject( idT, idsolidColorLayer, desc24 );
        executeAction( idsetd, desc23, DialogModes.NO );    
      eos
      do_javascript(color_js)  
    end  
  end    

end 

class OXY
  attr_accessor :should_pass_to_ps        
  
  # If an option is in the hash it means we have to use PS scripting to change that value.
  # Don't you just love the Applescript API for PS? //sarcasm 
  @should_pass_to_ps = { :text => [ 'color' ]}
end

@oxy = OXY.new

class OXYText
  
  @options = ActiveSupport::OrderedHash.new                                               
  @path = { :path => [], :index_path => [] } 
  @ps_obj = nil                         
  
  # Whether or not this has been created in Photoshop yet.
  @created_in_ps = false
           
  def initialize(name, contents, kind = OSA::AdobePhotoshopCS5::E580::PARAGRAPH_TEXT, path = [], index_path = [])
    @name = name 
    
    @options[:kind]     = kind  
    @options[:contents] = contents    
    
    @path[:path]       = path  
    @path[:index_path] = index_path  
    
    send_to_ps 
  end  
   
## Setters     
   
  def color=(rgb_color_hash)
    @options[:color] = rgb_color_hash     
  end        
  
  # Make Sure You Pass a OSA::AdobePhotoshopCS5::E580:: constant to this 
  # e.g kind = OSA::AdobePhotoshopCS5::E580::PARAGRAPH_TEXT
  def kind=(kind)
    @options[:kind] = kind 
  end 
  
  def contents=(text)
    @options[:contents] = text
  end
  
  # Sends the Layer To Photoshop
  def send_to_ps
    @ps_obj = @ps.add_layer(@name, 'text', @path)
    textobj  = @ps_obj.text_object   
    @options.each do |option, value|  
      if @oxy.should_pass_to_ps[:text].include?(option)  
        eval "@ps.text_#{option}_active(value)"
      else
        eval "textobj.#{option} = value"
      end  
    end
  end  
  
  def update     
    @ps.set_active(@ps_obj.name, @path)
    @options.each do |option, value|  
      if @oxy.should_pass_to_ps[:text].include?(option)  
        eval "@ps.text_#{option}_active(value)"
      else
        eval "textobj.#{option} = value"
      end  
    end
  end 
end  

class OXYParser
  def loop_elements
    @elems.each_with_index do |elem, index|
    end
  end
end