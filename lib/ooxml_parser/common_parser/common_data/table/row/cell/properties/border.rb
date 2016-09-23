require_relative 'table_cell_line'
module OoxmlParser
  # Border Data
  class Border < OOXMLDocumentObject
    attr_accessor :style, :color

    # Parse Border object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Border] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'style'
          @style = value.value.to_sym
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'color'
          @color = Color.parse_color_tag(node_child)
        end
      end
      self
    end
  end
end
