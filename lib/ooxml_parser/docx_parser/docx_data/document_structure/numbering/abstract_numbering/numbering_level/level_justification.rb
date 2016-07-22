module OoxmlParser
  # Class for storing Level Justification
  class LevelJustification
    # @return [String] value of start
    attr_accessor :value

    # Parse LevelJustification
    # @param [Nokogiri::XML:Node] node with LevelJustification
    # @return [LevelJustification] result of parsing
    def self.parse(node)
      level_jc = LevelJustification.new

      node.attributes.each do |key, value|
        case key
        when 'val'
          level_jc.value = value.value.to_s
        end
      end
      level_jc
    end
  end
end
