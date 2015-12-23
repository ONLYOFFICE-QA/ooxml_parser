module OoxmlParser
  class FrameProperties
    attr_accessor :drop_cap, :lines, :wrap, :vertical_anchor, :horizontal_anchor, :width, :height, :height_rule,
                  :horizontal_align, :vertical_align, :anchor_lock, :vertical_space, :horizontal_space,
                  :horizontal_position, :vertical_position

    def self.parse(frame_pr_node)
      frame_properties = FrameProperties.new
      frame_pr_node.attributes.each do |key, value|
        case key
        when 'dropCap'
          frame_properties.drop_cap = value.value.to_sym
        when 'lines'
          frame_properties.lines = value.value.to_i
        when 'wrap'
          frame_properties.lines = value.value.to_sym
        when 'vAnchor'
          frame_properties.vertical_anchor = value.value.to_sym
        when 'hAnchor'
          frame_properties.horizontal_anchor = value.value.to_sym
        when 'w'
          frame_properties.width = (value.value.to_f / 566.9).round(2)
        when 'h'
          frame_properties.height = (value.value.to_f / 566.9).round(2)
        when 'hRule'
          frame_properties.height_rule = value.value.to_s.sub('atLeast', 'at_least').to_sym
        when 'xAlign'
          frame_properties.horizontal_align = value.value.to_sym
        when 'yAlign'
          frame_properties.vertical_align = value.value.to_sym
        when 'anchorLock'
          frame_properties.anchor_lock = OOXMLDocumentObject.option_enabled?(frame_pr_node, anchorLock)
        when 'vSpace'
          frame_properties.vertical_space = (value.value.to_f / 566.9).round(2)
        when 'hSpace'
          frame_properties.horizontal_space = (value.value.to_f / 566.9).round(2)
        when 'x'
          frame_properties.horizontal_position = (value.value.to_f / 566.9).round(2)
        when 'y'
          frame_properties.vertical_position = (value.value.to_f / 566.9).round(2)
        end
      end
      frame_properties
    end
  end
end
