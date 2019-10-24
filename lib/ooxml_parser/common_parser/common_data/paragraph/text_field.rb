# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `fld` tags
  class TextField < OOXMLDocumentObject
    attr_accessor :id, :type, :text

    # Parse TextField object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextField] result of parsing
    def parse(node)
      @id = node.attribute('id').value
      @type = node.attribute('type').value
      node.xpath('*').each do |text_field_node_child|
        case text_field_node_child.name
        when 't'
          @text = text_field_node_child.text
        end
      end
      self
    end
  end
end
