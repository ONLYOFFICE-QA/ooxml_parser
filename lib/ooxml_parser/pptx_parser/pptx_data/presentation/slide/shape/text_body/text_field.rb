module OoxmlParser
  class TextField
    attr_accessor :id, :type, :text

    def self.parse(text_field_node)
      text_field = TextField.new
      text_field.id = text_field_node.attribute('id').value
      text_field.type = text_field_node.attribute('type').value
      text_field_node.xpath('*').each do |text_field_node_child|
        case text_field_node_child.name
        when 't'
          text_field.text = text_field_node_child.text
        end
      end
      text_field
    end
  end
end
