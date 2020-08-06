# frozen_string_literal: true

module OoxmlParser
  # Chart Label Properties, parse tag `dLbls`
  class DisplayLabelsProperties < OOXMLDocumentObject
    attr_accessor :position, :show_legend_key, :show_values, :show_x_axis_name, :show_y_axis_name
    # @return [True, False] is category name shown
    attr_accessor :show_category_name
    # @return [True, False] is series name shown
    attr_accessor :show_series_name
    # @return [True, False] is label is deleted
    attr_reader :delete

    def initialize(params = {})
      @show_legend_key = params.fetch(:show_legend_key, false)
      @show_values = params.fetch(:show_values, false)
      @parent = params[:parent]
    end

    # Parse DisplayLabelsProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DisplayLabelsProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'dLblPos'
          @position = value_to_symbol(node_child.attribute('val'))
        when 'showLegendKey'
          @show_legend_key = true if node_child.attribute('val').value == '1'
        when 'showVal'
          @show_values = true if node_child.attribute('val').value == '1'
        when 'showCatName'
          @show_category_name = option_enabled?(node_child)
        when 'showSerName'
          @show_series_name = option_enabled?(node_child)
        when 'delete'
          @delete = option_enabled?(node_child)
        end
      end
      self
    end
  end
end
