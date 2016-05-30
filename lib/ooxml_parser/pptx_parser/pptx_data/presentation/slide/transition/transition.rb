require_relative 'transition/sound_action'
require_relative 'transition_properties/transition_properties'
module OoxmlParser
  class Transition < OOXMLDocumentObject
    attr_accessor :speed, :properties, :sound_action, :advance_on_click, :delay, :duration

    def initialize(properties = TransitionProperties.new)
      @properties = properties
    end

    def self.parse(transition_node)
      transition = Transition.new
      transition_node.xpath('*').each do |transition_node_child|
        transition.properties = TransitionProperties.parse(transition_node_child)
        transition.sound_action = SoundAction.parse(transition_node_child)
      end
      transition_node.attributes.each do |key, value|
        case key
        when 'spd'
          transition.speed = value.value.to_sym
        when 'advClick'
          transition.advance_on_click = OOXMLDocumentObject.option_enabled?(transition_node, key)
        when 'advTm'
          transition.delay = value.value.to_f / 1_000.0
        when 'dur'
          transition.duration = value.value.to_f / 1_000.0
        end
      end
      transition
    end
  end
end
