module OoxmlParser
  # Class for parsing `hlinkClick`, `hyperlink` tags
  class Hyperlink < OOXMLDocumentObject
    attr_accessor :url, :tooltip, :coordinates, :id, :highlight_click, :action

    def initialize(link = nil,
                   tooltip = nil,
                   coordinates = nil,
                   parent: nil)
      @url = link
      @tooltip = tooltip
      @coordinates = coordinates
      @parent = parent
    end

    alias link url
    alias link_to url
    alias link= url=
    alias link_to= url=

    # Parse Hyperlink object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Hyperlink] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'location'
          @url = Coordinates.parse_coordinates_from_string(value.value)
        when 'id'
          @url = OOXMLDocumentObject.get_link_from_rels(node.attribute('id').value)
        when 'tooltip'
          @tooltip = value.value
        when 'ref'
          @coordinates = Coordinates.parse_coordinates_from_string(value.value)
        end
      end
      action = node.attribute('action').value unless node.attribute('action').nil?
      case action
      when 'ppaction://hlinkshowjump?jump=previousslide'
        @action = :previous_slide
      when 'ppaction://hlinkshowjump?jump=nextslide'
        @action = :next_slide
      when 'ppaction://hlinkshowjump?jump=firstslide'
        @action = :first_slide
      when 'ppaction://hlinkshowjump?jump=lastslide'
        @action = :last_slide
      when 'ppaction://hlinksldjump'
        @action = :slide
        @url = OOXMLDocumentObject.get_link_from_rels(node.attribute('id').value).scan(/\d+/).join('').to_i
      else
        unless node.attribute('id').nil?
          @action = :external_link
          @url = OOXMLDocumentObject.get_link_from_rels(node.attribute('id').value)
        end
      end
      return self unless node.attribute('highlightClick')
      @highlight_click = if node.attribute('highlightClick').value == '1'
                           true
                         else
                           false
                         end
      self
    end
  end
end
