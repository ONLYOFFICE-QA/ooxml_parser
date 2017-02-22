module OoxmlParser
  # Class for parsing `hlinkClick`, `hyperlink` tags
  class Hyperlink < OOXMLDocumentObject
    attr_accessor :url, :tooltip, :coordinates, :id, :highlight_click, :action
    attr_accessor :action_link
    attr_accessor :id

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
          @id = value.value
          @url = OOXMLDocumentObject.get_link_from_rels(@id)
        when 'tooltip'
          @tooltip = value.value
        when 'ref'
          @coordinates = Coordinates.parse_coordinates_from_string(value.value)
        when 'action'
          @action_link = value.value
        when 'highlightClick'
          @highlight_click = attribute_enabled?(value)
        end
      end
      case @action_link
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
        @url = OOXMLDocumentObject.get_link_from_rels(@id).scan(/\d+/).join('').to_i
      else
        unless @id.nil?
          @action = :external_link
          @url = OOXMLDocumentObject.get_link_from_rels(@id)
        end
      end
      self
    end
  end
end
