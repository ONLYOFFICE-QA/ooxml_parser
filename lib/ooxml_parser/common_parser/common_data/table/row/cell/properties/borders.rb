require_relative 'border'
module OoxmlParser
  class Borders < ParagraphBorders
    attr_accessor :left, :right, :top, :bottom, :inner_vertical, :inner_horizontal, :display, :between, :bar,
                  :top_left_to_bottom_right, :top_right_to_bottom_left, :offset_from

    def initialize(left = BordersProperties.new,
                   right = BordersProperties.new,
                   top = BordersProperties.new,
                   bottom = BordersProperties.new,
                   between = BordersProperties.new,
                   parent: nil)
      @left = left
      @right = right
      @top = top
      @bottom = bottom
      @between = between
      @parent = parent
    end

    def copy
      new_borders = Borders.new(@left, @right, @top, @bottom)
      new_borders.inner_vertical = @inner_vertical unless @inner_vertical.nil?
      new_borders.inner_horizontal = @inner_horizontal unless @inner_horizontal.nil?
      new_borders.between = @between unless @between.nil?
      new_borders.display = @display unless @display.nil?
      new_borders.bar = @bar unless bar.nil?
      new_borders
    end

    def ==(other)
      @left == other.left && @right == other.right && @top == other.top && @bottom == other.bottom if other.is_a?(Borders)
    end

    def each_border
      yield(bottom)
      yield(inner_horizontal)
      yield(inner_vertical)
      yield(left)
      yield(right)
      yield(top)
    end

    def each_side
      yield(bottom)
      yield(left)
      yield(right)
      yield(top)
    end

    def to_s
      "Left border: #{left}, Right: #{right}, Top: #{top}, Bottom: #{bottom}"
    end

    def visible?
      visible = false
      each_side do |current_size|
        visible ||= current_size.visible?
      end
      visible
    end

    # Parse Borders object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Borders] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'lnL', 'left'
          @left = TableCellLine.new(parent: self).parse(node_child)
        when 'lnR', 'right'
          @right = TableCellLine.new(parent: self).parse(node_child)
        when 'lnT', 'top'
          @top = TableCellLine.new(parent: self).parse(node_child)
        when 'lnB', 'bottom'
          @bottom = TableCellLine.new(parent: self).parse(node_child)
        when 'lnTlToBr', 'tl2br'
          @top_left_to_bottom_right = TableCellLine.new(parent: self).parse(node_child)
        when 'lnBlToTr', 'tr2bl'
          @top_right_to_bottom_left = TableCellLine.new(parent: self).parse(node_child)
        when 'insideV'
          @inner_vertical = TableCellLine.new(parent: self).parse(node_child)
        when 'insideH'
          @inner_horizontal = TableCellLine.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
