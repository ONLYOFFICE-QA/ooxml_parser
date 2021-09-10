# frozen_string_literal: true

require_relative 'icon_set/conditional_formatting_icon'
module OoxmlParser
  # Class for `iconSet` data
  class IconSet < OOXMLDocumentObject
    # @return [String] Name of icon set
    attr_reader :set
    # @return [Symbol] Specifies whether icons are shown in reverse order
    attr_reader :reverse
    # @return [Symbol] Specifies whether value is shown in a cell
    attr_reader :show_value
    # @return [Symbol] Specifies whether icon set is custom
    attr_reader :custom
    # @return [Array, ConditionalFormatValueObject] list of values
    attr_reader :values
    # @return [Array, ConditionalFormattingIcon] list of icons for custom sets
    attr_reader :icons

    def initialize(parent: nil)
      @values = []
      @icons = []
      @show_value = true
      super
    end

    # Parse IconSet data
    # @param [Nokogiri::XML:Element] node with IconSet data
    # @return [IconSet] value of IconSet data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'iconSet'
          @set = value.value.to_s
        when 'reverse'
          @reverse = attribute_enabled?(value)
        when 'showValue'
          @show_value = attribute_enabled?(value)
        when 'custom'
          @custom = attribute_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cfvo'
          @values << ConditionalFormatValueObject.new(parent: self).parse(node_child)
        when 'cfIcon'
          @icons << ConditionalFormattingIcon.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
