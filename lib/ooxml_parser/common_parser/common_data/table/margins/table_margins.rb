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
      super(parent: parent)
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
      "Default: #{is_default} top: #{@top}, bottom: #{@bottom}, left: #{@left}, right: #{@right}"
    end

    # Parse TableMargins object
    # @param margin_node [Nokogiri::XML:Element] node to parse
    # @return [TableMargins] result of parsing
    def parse(margin_node)
      margin_node.xpath('*').each do |cell_margin_node|
        case cell_margin_node.name
        when 'left'
          @left = OoxmlSize.new(parent: self).parse(cell_margin_node)
        when 'top'
          @top = OoxmlSize.new(parent: self).parse(cell_margin_node)
        when 'right'
          @right = OoxmlSize.new(parent: self).parse(cell_margin_node)
        when 'bottom'
          @bottom = OoxmlSize.new(parent: self).parse(cell_margin_node)
        end
      end
      self
    end
  end
end
