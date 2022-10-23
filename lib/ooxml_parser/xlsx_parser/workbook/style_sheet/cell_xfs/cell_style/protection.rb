# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <protection> tag
  class Protection < OOXMLDocumentObject
    # @return [True, False] Specifies if cell is locked
    attr_reader :locked
    # @return [True, False] Specifies if formulas in cell are hidden
    attr_reader :hidden

    def initialize(parent: nil)
      @locked = true
      @hidden = false
      super
    end

    # Parse Protection data
    # @param [Nokogiri::XML:Element] node with Protection data
    # @return [Sheet] value of Protection
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'locked'
          @locked = boolean_attribute_value(value.value)
        when 'hidden'
          @hidden = boolean_attribute_value(value.value)
        end
      end
      self
    end
  end
end
