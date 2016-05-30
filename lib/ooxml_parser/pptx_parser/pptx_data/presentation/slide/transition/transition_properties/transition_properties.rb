module OoxmlParser
  class TransitionProperties
    attr_accessor :type, :through_black, :direction, :orientation, :spokes

    def initialize(type = nil)
      @type = type
    end

    # Parse TransitionProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TransitionProperties] result of parsing
    def self.parse(node)
      transition_properties = TransitionProperties.new
      transition_properties.type = node.name.to_sym
      case node.name
      when 'blinds', 'checker', 'comb', 'cover', 'pull', 'push', 'randomBar', 'strips', 'wipe', 'zoom', 'warp'
        transition_properties.direction = Alignment.parse(node.attribute('dir')) if node.attribute('dir')
      when 'cut', 'fade'
        transition_properties.through_black = OOXMLDocumentObject.option_enabled?(node, 'thruBlk')
      when 'split'
        transition_properties.direction = Alignment.parse(node.attribute('dir')) if node.attribute('dir')
        transition_properties.orientation = node.attribute('orient').value.to_sym if node.attribute('orient')
      when 'wheel', 'wheelReverse'
        transition_properties.spokes = OOXMLDocumentObject.option_enabled?(node, 'spokes')
      end
      transition_properties
    end
  end
end
