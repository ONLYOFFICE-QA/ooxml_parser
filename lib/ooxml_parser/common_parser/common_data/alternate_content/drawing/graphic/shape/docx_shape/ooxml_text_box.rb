module OoxmlParser
  class OOXMLTextBox < OOXMLDocumentObject
    attr_accessor :properties, :elements

    def initialize(properties = nil, elements = [])
      @properties = properties
      @elements = elements
    end

    def self.parse(text_body_node, parent: nil)
      text_body = OOXMLTextBox.new
      text_body.parent = parent
      text_box_content_node = text_body_node.xpath('w:txbxContent').first
      text_box_content_node.xpath('*').each_with_index do |text_body_node_child, index|
        case text_body_node_child.name
        when 'p'
          text_body.elements << DocxParagraph.new.parse(text_body_node_child, index, parent: text_body)
        when 'tbl'
          text_body.elements << Table.new(parent: text_body).parse(text_body_node_child, index)
        when 'bodyPr'
          text_body.properties = OOXMLShapeBodyProperties.new(parent: text_body).parse(text_body_node_child)
        end
      end
      text_body
    end
  end
end
