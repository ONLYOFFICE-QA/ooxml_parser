module OoxmlParser
  class FrameProperties < OOXMLDocumentObject
    attr_accessor :drop_cap, :lines, :wrap, :vertical_anchor, :horizontal_anchor, :width, :height, :height_rule,
                  :horizontal_align, :vertical_align, :anchor_lock, :vertical_space, :horizontal_space,
                  :horizontal_position, :vertical_position

    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'dropCap'
          @drop_cap = value.value.to_sym
        when 'lines'
          @lines = value.value.to_i
        when 'wrap'
          @lines = value.value.to_sym
        when 'vAnchor'
          @vertical_anchor = value.value.to_sym
        when 'hAnchor'
          @horizontal_anchor = value.value.to_sym
        when 'w'
          @width = OoxmlSize.new(value.value.to_f)
        when 'h'
          @height = OoxmlSize.new(value.value.to_f)
        when 'hRule'
          @height_rule = value.value.to_s.sub('atLeast', 'at_least').to_sym
        when 'xAlign'
          @horizontal_align = value.value.to_sym
        when 'yAlign'
          @vertical_align = value.value.to_sym
        when 'anchorLock'
          @anchor_lock = attribute_enabled?(value)
        when 'vSpace'
          @vertical_space = OoxmlSize.new(value.value.to_f)
        when 'hSpace'
          @horizontal_space = OoxmlSize.new(value.value.to_f)
        when 'x'
          @horizontal_position = OoxmlSize.new(value.value.to_f)
        when 'y'
          @vertical_position = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end
