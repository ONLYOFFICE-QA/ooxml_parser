# frozen_string_literal: true

module OoxmlParser
  # Class for working with Table Margins
  class TableMargins < OOXMLDocumentObject
    attr_accessor :is_default, :top, :bottom, :left, :right

    def initialize(is_default = true, top = nil, bottom = nil, left = nil, right = nil, parent: nil)
      @is_default = is_default
      @top = top
      @bottom = bottom
      @left = left
      @right = right
      @parent = parent
    end

    # TODO: Separate @is_default attribute and remove this method
    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      instance_variables.each do |current_attribute|
        next if current_attribute == :@parent
        next if current_attribute == :@is_default
        return false unless instance_variable_get(current_attribute) == other.instance_variable_get(current_attribute)
      end
      true
    end

    # @return [String] result of convert of object to string
    def to_s
      'Default: ' + is_default.to_s + ' top: ' + @top.to_s + ', bottom: ' + @bottom.to_s + ', left: ' + @left.to_s + ', right: ' + @right.to_s
    end

    def parse(margin_node)
      margin_node.xpath('*').each do |cell_margin_node|
        case cell_margin_node.name
        when 'left'
          @left = OoxmlSize.new(cell_margin_node.attribute('w').value.to_f)
        when 'top'
          @top = OoxmlSize.new(cell_margin_node.attribute('w').value.to_f)
        when 'right'
          @right = OoxmlSize.new(cell_margin_node.attribute('w').value.to_f)
        when 'bottom'
          @bottom = OoxmlSize.new(cell_margin_node.attribute('w').value.to_f)
        end
      end
      self
    end
  end
end
