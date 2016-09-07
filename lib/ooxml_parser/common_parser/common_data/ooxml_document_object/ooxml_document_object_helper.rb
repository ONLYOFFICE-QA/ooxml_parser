# frozen_string_literal: true
# Module for helper methods for OOXMLDocumentObject
module OoxmlDocumentObjectHelper
  VALUE_TO_SYMBOL_HASH = { l: :left,
                           ctr: :center,
                           r: :right,
                           just: :justify,
                           b: :bottom,
                           t: :top,
                           tr: :top_right,
                           tl: :top_left,
                           br: :bottom_right,
                           bl: :bottom_left,
                           dist: :distributed,
                           tb: :horizontal,
                           rl: :rotate_on_90,
                           lr: :rotate_on_270,
                           inset: :in,
                           lu: :left_up,
                           ru: :right_up,
                           ld: :left_down,
                           rd: :right_down,
                           d: :down,
                           u: :up,
                           rightMargin: :right_margin,
                           leftMargin: :left_margin,
                           bottomMargin: :bottom_margin,
                           topMargin: :top_margin,
                           undOvr: :under_over,
                           subSup: :subscript_superscript,
                           lrTb: :left_to_right_top_to_bottom,
                           tbRl: :top_to_bottom_right_to_left,
                           btLr: :bottom_to_top_left_to_right,
                           singlelevel: :single_level,
                           multilevel: :multi_level,
                           hybridMultilevel: :hybrid_multi_level,
                           rnd: :round,
                           sq: :square }.freeze
  # Convert value to human readable symbol
  # @param [String] value to convert
  # @return [Symbol]
  def value_to_symbol(value)
    symbol = VALUE_TO_SYMBOL_HASH[value.value.to_sym]
    return value.value.to_sym if symbol.nil?
    symbol
  end
end
