# Class for describing Alignment type
module OoxmlParser
  class Alignment
    # Parse TransitionDirection
    # @param [Nokogiri::XML:Element] node with Alignment
    # @return [Symbol] value of Alignment
    def self.parse(node)
      case node.value
      when 'l'
        :left
      when 'ctr'
        :center
      when 'r'
        :right
      when 'just'
        :justify
      when 'b'
        :bottom
      when 't'
        :top
      when 'tr'
        :top_right
      when 'tl'
        :top_left
      when 'br'
        :bottom_right
      when 'bl'
        :bottom_left
      when 'dist'
        :distributed
      when 'tb'
        :horizontal
      when 'rl'
        :rotate_on_90
      when 'lr'
        :rotate_on_270
      when 'inset'
        :in
      when 'lu'
        :left_up
      when 'ru'
        :right_up
      when 'ld'
        :left_down
      when 'rd'
        :right_down
      when 'd'
        :down
      when 'u'
        :up
      when 'rightMargin'
        :right_margin
      when 'leftMargin'
        :left_margin
      when 'bottomMargin'
        :bottom_margin
      when 'topMargin'
        :top_margin
      when 'undOvr'
        :under_over
      when 'subSup'
        :subscript_superscript
      else
        node.value.to_sym
      end
    end
  end
end
