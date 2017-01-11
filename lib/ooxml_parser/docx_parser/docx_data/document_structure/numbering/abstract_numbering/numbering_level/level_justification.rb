module OoxmlParser
  # Class for storing Level Justification, `lvlJc` tag
  class LevelJustification < OOXMLDocumentObject
    # @return [String] value of start
    attr_accessor :value

    # Parse LevelJustification
    # @param [Nokogiri::XML:Node] node with LevelJustification
    # @return [LevelJustification] result of parsing
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
