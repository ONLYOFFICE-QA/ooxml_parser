# frozen_string_literal: true

module OoxmlParser
  # Class for describing Shape Guide
  class ShapeGuide < OOXMLDocumentObject
    # @return [String] name of guide
    attr_accessor :name
    # @return [String] shape guide formula
    attr_accessor :formula

    # Parse ShapeGuide
    # @param [Nokogiri::XML:Node] node with Shape Guide
    # @return [ShapeGuide] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_sym
        when 'fmla'
          @formula = value.value
        end
      end
      self
    end
  end
end
