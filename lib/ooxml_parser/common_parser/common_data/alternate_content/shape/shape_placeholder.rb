module OoxmlParser
  class ShapePlaceholder < OOXMLDocumentObject
    attr_accessor :type, :has_custom_prompt

    def self.parse(placeholder_node)
      placeholder = ShapePlaceholder.new
      placeholder_node.attributes.each do |key, value|
        case key
        when 'type'
          placeholder.type = value.value.to_sym
        when 'hasCustomPrompt'
          placeholder.has_custom_prompt = OOXMLDocumentObject.option_enabled?(placeholder_node, 'hasCustomPrompt')
        end
      end
      placeholder
    end
  end
end
