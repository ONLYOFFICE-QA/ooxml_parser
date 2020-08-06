# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `txbx` tags
  class OOXMLTextBox < OOXMLDocumentObject
    attr_accessor :elements

    def initialize(parent: nil)
      @elements = []
      super
    end

    # Parse OOXMLTextBox object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OOXMLTextBox] result of parsing
    def parse(node)
      text_box_content_node = node.xpath('w:txbxContent').first
      text_box_content_node.xpath('*').each_with_index do |node_child, index|
        case node_child.name
        when 'p'
          @elements << DocxParagraph.new(parent: self).parse(node_child, index)
        when 'tbl'
          @elements << Table.new(parent: self).parse(node_child, index)
        end
      end
      self
    end
  end
end
