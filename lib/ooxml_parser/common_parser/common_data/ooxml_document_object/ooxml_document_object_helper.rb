# Module for helper methods for OOXMLDocumentObject
module OoxmlDocumentObjectHelper
  # Convert object to hash
  # @return [Hash]
  def to_hash
    result_hash = {}
    instance_variables.each do |current_attribute|
      next if current_attribute == :@parent
      attribute_value = instance_variable_get(current_attribute)
      next unless attribute_value
      if attribute_value.is_a?(Array)
        attribute_value.each_with_index do |object_element, index|
          result_hash["#{current_attribute}_#{index}".to_sym] = object_element.to_hash
        end
      else
        result_hash[current_attribute.to_sym] = if attribute_value.respond_to?(:to_hash)
                                                  attribute_value.to_hash
                                                else
                                                  attribute_value.to_s
                                                end
      end
    end
    result_hash
  end

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
                           contentLocked: :content_locked,
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
                           beneathText: :beneath_text,
                           greaterThan: :greater_than,
                           pageBottom: :page_bottom,
                           sectEnd: :section_end,
                           docEnd: :document_end,
                           eachSect: :each_section,
                           eachPage: :each_page,
                           sngStrike: :single,
                           dblStrike: :double,
                           noStrike: :none,
                           lgDash: :large_dash,
                           dashDot: :dash_dot,
                           dashDotDot: :dash_dot_dot,
                           mediumDashed: :medium_dashed,
                           mediumDashDot: :medium_dash_dot,
                           mediumDashDotDot: :medium_dash_dot_dot,
                           lgDashDot: :large_dash_dot,
                           lgDashDotDot: :large_dash_dot_dot,
                           sdtLocked: :sdt_locked,
                           sdtContentLocked: :sdt_content_locked,
                           sysDash: :system_dash,
                           sysDot: :system_dot,
                           sysDashDot: :system_dash_dot,
                           sysDashDotDot: :system_dash_dot_dot,
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
                           maxMin: :max_min,
                           middleDot: :middle_dot,
                           minMax: :min_max,
                           multilevel: :multi_level,
                           nextTo: :next_to,
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
    !(status == 'false' || status == 'off' || status == '0')
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
