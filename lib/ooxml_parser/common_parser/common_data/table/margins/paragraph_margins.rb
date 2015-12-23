module OoxmlParser
  class ParagraphMargins < TableMargins
    def initialize(top = 0.0, bottom = 0.0, left = 0.0, right = 0.0)
      @top = top
      @bottom = bottom
      @left = left
      @right = right
    end

    def round(count_of_digits = 1)
      top = @top.round(count_of_digits)
      bottom = @bottom.round(count_of_digits)
      left = @left.round(count_of_digits)
      right = @right.round(count_of_digits)
      ParagraphMargins.new(top, bottom, left, right)
    end

    def self.parse(text_body_props_node)
      margins = ParagraphMargins.new(0.127, 0.127, 0.254, 0.254)
      text_body_props_node.attributes.each do |key, value|
        case key
        when 'bIns', 'marB'
          margins.bottom = (value.value.to_f / 360_000.0).round(3)
        when 'tIns', 'marT'
          margins.top = (value.value.to_f / 360_000.0).round(3)
        when 'lIns', 'marL'
          margins.left = (value.value.to_f / 360_000.0).round(3)
        when 'rIns', 'marR'
          margins.right = (value.value.to_f / 360_000.0).round(3)
        end
      end
      margins
    end
  end
end
