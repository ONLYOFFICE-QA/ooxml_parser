# frozen_string_literal: true

module OoxmlParser
  # Class for working with ParagraphMargins
  class ParagraphMargins < TableMargins
    def initialize(top = OoxmlSize.new(0),
                   bottom = OoxmlSize.new(0),
                   left = OoxmlSize.new(0),
                   right = OoxmlSize.new(0),
                   parent: nil)
      @top = top
      @bottom = bottom
      @left = left
      @right = right
      @parent = parent
    end

    def parse(text_body_props_node)
      text_body_props_node.attributes.each do |key, value|
        case key
        when 'bIns', 'marB'
          @bottom = OoxmlSize.new(value.value.to_f, :emu)
        when 'tIns', 'marT'
          @top = OoxmlSize.new(value.value.to_f, :emu)
        when 'lIns', 'marL'
          @left = OoxmlSize.new(value.value.to_f, :emu)
        when 'rIns', 'marR'
          @right = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      self
    end
  end
end
