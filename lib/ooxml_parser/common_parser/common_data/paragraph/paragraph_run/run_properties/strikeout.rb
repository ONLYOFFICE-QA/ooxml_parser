# Class for describing strikeout
module OoxmlParser
  class Strikeout
    # Parse Strikeout
    # @param [Nokogiri::XML:Attr] node with Strikeout
    # @return [Symbol] value of Strikeout
    def self.parse(node)
      case node.value
      when 'sngStrike'
        :single
      when 'dblStrike'
        :double
      end
    end
  end
end
