module OoxmlParser
  class AnimationEffect < OOXMLDocumentObject
    attr_accessor :transition, :filter, :behavior

    # Parse AnimationEffect object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [AnimationEffect] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'transition'
          @transition = value.value
        when 'filter'
          @filter = value.value
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cBhvr'
          @behavior = Behavior.parse(node_child)
        end
      end
      self
    end
  end
end
