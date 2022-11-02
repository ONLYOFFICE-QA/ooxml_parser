# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:pPrDefault` tags
  class ParagraphPropertiesDefault < OOXMLDocumentObject
    # @return [ParagraphProperties] properties of run
    attr_accessor :paragraph_properties
    # @return [Nokogiri::XML:Element] raw node of tag
    attr_reader :raw_node

    # Parse ParagraphPropertiesDefault object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphPropertiesDefault] result of parsing
    def parse(node)
      @raw_node = node
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pPr'
          @paragraph_properties = ParagraphProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
