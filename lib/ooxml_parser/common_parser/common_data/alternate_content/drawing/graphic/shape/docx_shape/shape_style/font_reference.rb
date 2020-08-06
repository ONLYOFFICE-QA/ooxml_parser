# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `a:fontRef` tags
  class FontReference < OOXMLDocumentObject
    # @return [String] Font Collection Index
    attr_reader :index
    # @return [Color] scheme color of FontReference
    attr_reader :scheme_color

    # Parse FontReference object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FontReference] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'idx'
          @index = value.value.to_s
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
