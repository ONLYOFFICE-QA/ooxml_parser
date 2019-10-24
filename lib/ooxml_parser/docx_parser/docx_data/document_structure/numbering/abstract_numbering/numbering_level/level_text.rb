# frozen_string_literal: true

module OoxmlParser
  # Class for storing Level Text, `lvlText` tag
  class LevelText < OOXMLDocumentObject
    # @return [String] value of start
    attr_accessor :value

    # Parse LevelText
    # @param [Nokogiri::XML:Node] node with LevelText
    # @return [LevelText] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_s
        end
      end
      self
    end
  end
end
