# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `formPr` tag
  class FormProperties < OOXMLDocumentObject
    # @return [String] field key
    attr_reader :key
    # @return [String] text of tooltip
    attr_reader :help_text
    # @return [True, False] specifies if field is required
    attr_reader :required
    # @return [Shade] shade of field
    attr_reader :shade
    # @return [BordersProperties] border of field
    attr_reader :border

    def initialize(parent: nil)
      @required = false
      super
    end

    # Parse FormProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FormProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'key'
          @key = value.value
        when 'helpText'
          @help_text = value.value
        when 'required'
          @required = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'shd'
          @shade = Shade.new(parent: self).parse(node_child)
        when 'border'
          @border = BordersProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
