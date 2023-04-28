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
          textbox_paragraph = DocxParagraph.new
          textbox_paragraph.spacing = Spacing.new(0, 0.35, 1.15, :multiple)
          elements << textbox_paragraph.parse(textbox_element, i, parent: parent)
        when 'tbl'
          elements << Table.new(parent: parent).parse(textbox_element, i)
        end
      end
      elements
    end
  end
end
