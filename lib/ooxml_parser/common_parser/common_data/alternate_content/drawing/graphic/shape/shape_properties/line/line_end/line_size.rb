# Class for describing Line Size
module OoxmlParser
  class LineSize
    # Parse LineSize
    # @param [Nokogiri::XML:Attr] node with LineSize
    # @return [Symbol] value of LineSize
    def self.parse(node)
      case node
      when 'lg'
        :large
      when 'med'
        :medium
      when 'sm'
        :small
      else
        node.to_sym
      end
    end
  end
end
