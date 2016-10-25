require_relative 'transition/sound_action'
require_relative 'transition_properties/transition_properties'
module OoxmlParser
  class Transition < OOXMLDocumentObject
    attr_accessor :speed, :properties, :sound_action, :advance_on_click, :delay, :duration

    def parse(node)
      node.xpath('*').each do |node_child|
        @properties = TransitionProperties.new(parent: self).parse(node_child)
        @sound_action = SoundAction.parse(node_child)
      end
      node.attributes.each do |key, value|
        case key
        when 'spd'
          @speed = value.value.to_sym
        when 'advClick'
          @advance_on_click = attribute_enabled?(value)
        when 'advTm'
          @delay = value.value.to_f / 1_000.0
        when 'dur'
          @duration = value.value.to_f / 1_000.0
        end
      end
      self
    end
  end
end
