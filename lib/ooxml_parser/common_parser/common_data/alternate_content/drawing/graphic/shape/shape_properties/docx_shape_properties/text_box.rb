# Class for working with TextBox (w:txbxContent)
module OoxmlParser
  class TextBox
    # Parse TextBox List
    # @param [Nokogiri::XML:Node] node with TextBox
    # @return [Array] array of elements
    def self.parse_list(node, parent: nil)
      elements = []
      text_box_content_node = node.xpath('w:txbxContent').first
      unless text_box_content_node.nil?
        text_box_content_node.xpath('*').each_with_index do |textbox_element, i|
          case textbox_element.name
          when 'p'
            DocumentStructure.default_paragraph_style = DocxParagraph.new
            DocumentStructure.default_paragraph_style.spacing = Spacing.new(0, 0.35, 1.15, :multiple)
            elements << DocumentStructure.default_paragraph_style.dup.parse(textbox_element, i, parent: parent)
          when 'tbl'
            elements << Table.new(parent: parent).parse(textbox_element, i)
          end
        end
      end
      elements
    end
  end
end
