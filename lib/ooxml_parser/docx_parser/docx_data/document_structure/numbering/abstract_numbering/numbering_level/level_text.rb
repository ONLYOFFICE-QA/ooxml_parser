module OoxmlParser
  # Class for storing Level Text
  class LevelText
    # @return [String] value of start
    attr_accessor :value

    # Parse LevelText
    # @param [Nokogiri::XML:Node] node with LevelText
    # @return [LevelText] result of parsing
    def self.parse(node)
      level_text = LevelText.new

      node.attributes.each do |key, value|
        case key
        when 'val'
          level_text.value = value.value.to_s
        end
      end
      level_text
    end
  end
end
