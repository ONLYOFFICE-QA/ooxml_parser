# frozen_string_literal: true

require_relative 'transition_properties/wheel'
module OoxmlParser
  class TransitionProperties < OOXMLDocumentObject
    attr_accessor :type, :through_black, :direction, :orientation, :spokes
    # @return [Wheel] Wheel Slide Transition
    attr_reader :wheel

    # Parse TransitionProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TransitionProperties] result of parsing
    def parse(node)
      @type = node.name.to_sym
      case node.name
      when 'blinds', 'checker', 'comb', 'cover', 'pull', 'push', 'randomBar', 'strips', 'wipe', 'zoom', 'warp'
        @direction = value_to_symbol(node.attribute('dir')) if node.attribute('dir')
      when 'cut', 'fade'
        @through_black = option_enabled?(node, 'thruBlk')
      when 'split'
        @direction = value_to_symbol(node.attribute('dir')) if node.attribute('dir')
        @orientation = node.attribute('orient').value.to_sym if node.attribute('orient')
      when 'wheel'
        @wheel = Wheel.new(parent: self).parse(node)
      end
      self
    end
  end
end
