require_relative 'sound_action/sound'
module OoxmlParser
  class SoundAction
    attr_accessor :start_sound, :end_sound

    def self.parse(sound_action_node)
      sound_action = SoundAction.new
      sound_action_node.xpath('*').each do |sound_action_node_child|
        case sound_action_node_child.name
        when 'stSnd'
          sound_action.start_sound = Sound.parse(sound_action_node_child.xpath('p:snd').first) unless sound_action_node_child.xpath('p:snd').first.nil?
        when 'endSnd'
          sound_action.end_sound = Sound.parse(sound_action_node_child.xpath('p:snd').first) unless sound_action_node_child.xpath('p:snd').first.nil?
        end
      end
      sound_action
    end
  end
end
