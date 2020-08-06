# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `a:fillRef` tags
  class FillReference < OOXMLDocumentObject
    # @return [Integer] Style Matrix Index
    attr_reader :index
    # @return [Color] scheme color of FillReference
    attr_reader :scheme_color

    # Parse FillReference object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FillReference] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'idx'
          @index = value.value.to_f
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'schemeClr'
          @scheme_color = Color.new(parent: self).parse_scheme_color(node_child)
        end
      end
      self
    end
  end
end
