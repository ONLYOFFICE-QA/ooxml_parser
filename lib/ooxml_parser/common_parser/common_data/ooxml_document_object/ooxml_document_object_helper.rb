# frozen_string_literal: true
# Module for helper methods for OOXMLDocumentObject
module OoxmlDocumentObjectHelper
  private

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
                           lg: :large,
                           med: :medium,
                           sm: :small,
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

  def option_enabled?(node, attribute_name = 'val')
    return true if node.attributes.empty?
    return true if node.to_s == '1'
    return false if node.to_s == '0'
    return false if node.attribute(attribute_name).nil?
    status = node.attribute(attribute_name).value
    status == 'true' || status == 'on' || status == '1'
  end

  def attribute_enabled?(node, attribute_name = 'val')
    return true if node.to_s == '1'
    return false if node.to_s == '0'
    return false if node.attribute(attribute_name).nil?
    status = node.attribute(attribute_name).value
    status == 'true' || status == 'on' || status == '1'
  end

  def root_object
    tree_object = self
    tree_object = tree_object.parent until tree_object.parent.nil?
    tree_object
  end
end
