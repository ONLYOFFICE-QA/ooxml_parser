# frozen_string_literal: true

module OoxmlParser
  # Class for `cfvo` data
  class ConditionalFormatValueObject < OOXMLDocumentObject
    # @return [Symbol] Specifies whether value is shown in a cell
    attr_reader :type
    # @return [Formula] formula
    attr_reader :formula

    # Parse ConditionalFormatValueObject data
    # @param [Nokogiri::XML:Element] node with ConditionalFormatValueObject data
    # @return [ConditionalFormatValueObject] value of ConditionalFormatValueObject data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'f'
          @formula = Formula.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
