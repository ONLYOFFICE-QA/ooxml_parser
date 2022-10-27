# frozen_string_literal: true

module OoxmlParser
  # Class for working with any tag contained only value
  class ValuedChild < OOXMLDocumentObject
    # @return [String] value of tag
    attr_accessor :value
    # @return [String] type of value
    attr_reader :type

    def initialize(type = :string, parent: nil)
      @type = type
      super(parent: parent)
    end

    # Parse ValuedChild object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ValuedChild] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_s if type == :string
          @value = value_to_symbol(value) if type == :symbol
          @value = value.value.to_i if type == :integer
          @value = value.value.to_f if type == :float
        end
      end
      self
    end
  end
end
