module OoxmlParser
  class Hyperlink < OOXMLDocumentObject
    attr_accessor :url, :tooltip, :coordinates, :id, :highlight_click, :action

    def initialize(link = nil, tooltip = nil, coordinates = nil)
      @url = link
      @tooltip = tooltip
      @coordinates = coordinates
    end

    alias link url
    alias link_to url
    alias link= url=
    alias link_to= url=

    def self.parse(on_click_hyperlink_node)
      hyperlink = Hyperlink.new
      on_click_hyperlink_node.attributes.each do |key, value|
        case key
        when 'location'
          hyperlink.url = Coordinates.parse_coordinates_from_string(value.value)
        when 'id'
          hyperlink.url = OOXMLDocumentObject.get_link_from_rels(on_click_hyperlink_node.attribute('id').value)
        when 'tooltip'
          hyperlink.tooltip = value.value
        when 'ref'
          hyperlink.coordinates = Coordinates.parse_coordinates_from_string(value.value)
        end
      end
      action = on_click_hyperlink_node.attribute('action').value unless on_click_hyperlink_node.attribute('action').nil?
      case action
      when 'ppaction://hlinkshowjump?jump=previousslide'
        hyperlink.action = :previous_slide
      when 'ppaction://hlinkshowjump?jump=nextslide'
        hyperlink.action = :next_slide
      when 'ppaction://hlinkshowjump?jump=firstslide'
        hyperlink.action = :first_slide
      when 'ppaction://hlinkshowjump?jump=lastslide'
        hyperlink.action = :last_slide
      when 'ppaction://hlinksldjump'
        hyperlink.action = :slide
        hyperlink.link_to = OOXMLDocumentObject.get_link_from_rels(on_click_hyperlink_node.attribute('id').value).scan(/\d+/).join('').to_i
      else
        unless on_click_hyperlink_node.attribute('id').nil?
          hyperlink.action = :external_link
          hyperlink.link_to = OOXMLDocumentObject.get_link_from_rels(on_click_hyperlink_node.attribute('id').value)
        end
      end
      hyperlink.highlight_click = StringHelper.to_bool(on_click_hyperlink_node.attribute('highlightClick').value) unless on_click_hyperlink_node.attribute('highlightClick').nil?
      hyperlink
    end
  end
end
