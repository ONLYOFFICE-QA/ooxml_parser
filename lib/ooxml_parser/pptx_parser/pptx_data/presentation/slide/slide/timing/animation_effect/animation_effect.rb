module OoxmlParser
  class AnimationEffect
    attr_accessor :transition, :filter, :behavior

    def initialize(behavior = nil)
      @behavior = behavior
    end

    def self.parse(animation_effect_node)
      animation_effect = AnimationEffect.new
      animation_effect.transition = animation_effect_node.attribute('transition').value
      animation_effect.filter = animation_effect_node.attribute('filter').value
      animation_effect_node.xpath('*').each do |animation_effect_node_child|
        case animation_effect_node_child.name
        when 'cBhvr'
          animation_effect.behavior = Behavior.parse(animation_effect_node_child)
        end
      end
      animation_effect
    end
  end
end
