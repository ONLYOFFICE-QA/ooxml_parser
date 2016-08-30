require_relative 'shape_placeholder'
module OoxmlParser
  # Class for parsing `nvPr` object
  class NonVisualProperties < OOXMLDocumentObject
    attr_accessor :placeholder, :is_photo, :user_drawn

    # Parse NonVisualProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [NonVisualProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'isPhoto'
          @is_photo = attribute_enabled?(value)
        when 'userDrawn'
          @user_drawn = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'ph'
          @placeholder = ShapePlaceholder.parse(node_child)
        end
      end
      self
    end
  end
end
