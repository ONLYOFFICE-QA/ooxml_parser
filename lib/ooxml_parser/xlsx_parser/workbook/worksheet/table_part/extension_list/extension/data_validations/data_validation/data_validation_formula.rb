# frozen_string_literal: true

module OoxmlParser
  # Class for `x14:formula1` or `x14:formula2` data
  class DataValidationFormula < OOXMLDocumentObject
    # @return [Formula] value of formula
    attr_reader :formula

    # Parse DataValidationFormula data
    # @param [Nokogiri::XML:Element] node with DataValidationFormula data
    # @return [DataValidationFormula] value of DataValidationFormula data
    def parse(node)
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
