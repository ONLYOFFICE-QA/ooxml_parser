# frozen_string_literal: true

module OoxmlParser
  # Class for working with TextBox (w:txbxContent)
  class TextBox < OOXMLDocumentObject
    # Parse TextBox List
    # @param [Nokogiri::XML:Node] node with TextBox
    # @return [Array] array of elements
    def parse_list(node)
      elements = []
      text_box_content_node = node.xpath('w:txbxContent').first

      text_box_content_node&.xpath('*')&.each_with_index do |textbox_element, i|
        case textbox_element.name
        when 'p'
          DocumentStructure.default_paragraph_style = DocxParagraph.new
          DocumentStructure.default_paragraph_style.spacing = Spacing.new(0, 0.35, 1.15, :multiple)
          elements << DocumentStructure.default_paragraph_style.dup.parse(textbox_element, i, parent: parent)
        when 'tbl'
          elements << Table.new(parent: parent).parse(textbox_element, i)
        end
      end
      elements
    end
  end
end
