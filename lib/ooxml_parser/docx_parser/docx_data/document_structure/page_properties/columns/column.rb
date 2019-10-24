# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:col` object
  class Column < OOXMLDocumentObject
    attr_accessor :width, :space, :separator

    # Parse Column
    # @param [Nokogiri::XML:Node] node with Column
    # @return [Column] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'w'
          @width = OoxmlSize.new(value.value.to_f)
        when 'space'
          @space = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end
