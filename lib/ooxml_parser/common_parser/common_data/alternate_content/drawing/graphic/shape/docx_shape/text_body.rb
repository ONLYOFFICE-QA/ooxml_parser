module OoxmlParser
  class TextBody
    attr_accessor :properties, :paragraphs

    def initialize(properties = nil, paragraphs = [])
      @properties = properties
      @paragraphs = paragraphs
    end

    def self.parse(text_body_node)
      text_body = TextBody.new
      text_body_node.xpath('*').each do |text_body_node_child|
        case text_body_node_child.name
        when 'p'
          text_body.paragraphs << Paragraph.parse(text_body_node_child)
        when 'bodyPr'
          text_body.properties = OOXMLShapeBodyProperties.new(parent: text_body).parse(text_body_node_child)
        end
      end
      text_body
    end
  end
end
