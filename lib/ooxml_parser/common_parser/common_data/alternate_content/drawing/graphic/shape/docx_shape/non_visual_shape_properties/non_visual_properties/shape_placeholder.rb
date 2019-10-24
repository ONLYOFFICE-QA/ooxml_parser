# frozen_string_literal: true

module OoxmlParser
  # Class for data about `ph` node
  class ShapePlaceholder < OOXMLDocumentObject
    attr_accessor :type, :has_custom_prompt

    # Parse ShapePlaceholder object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ShapePlaceholder] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        when 'hasCustomPrompt'
          @has_custom_prompt = attribute_enabled?(node)
        end
      end
      self
    end
  end
end
