require_relative 'border'
module OoxmlParser
  class Borders < ParagraphBorders
    attr_accessor :left, :right, :top, :bottom, :inner_vertical, :inner_horizontal, :display, :between, :bar,
                  :top_left_to_bottom_right, :top_right_to_bottom_left, :offset_from

    def initialize(left = BordersProperties.new,
                   right = BordersProperties.new,
                   top = BordersProperties.new,
                   bottom = BordersProperties.new,
                   between = BordersProperties.new)
      @left = left
      @right = right
      @top = top
      @bottom = bottom
      @between = between
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

    def self.parse(borders_node)
      borders = Borders.new
      borders_node.xpath('*').each do |border_node|
        case border_node.name
        when 'lnL', 'left'
          borders.left = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'lnR', 'right'
          borders.right = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'lnT', 'top'
          borders.top = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'lnB', 'bottom'
          borders.bottom = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'lnTlToBr', 'tl2br'
          borders.top_left_to_bottom_right = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'lnBlToTr', 'tr2bl'
          borders.top_right_to_bottom_left = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'insideV'
          borders.inner_vertical = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        when 'insideH'
          borders.inner_horizontal = TableCellLine.new(PresentationFill.parse(border_node), LineJoin.parse(border_node), parent: borders).parse(border_node)
        end
      end
      borders
    end

    def self.parse_from_style(style_number)
      border_style_node = XLSXWorkbook.styles_node.xpath('//xmlns:border')[style_number.to_i]
      cell_borders = Borders.new
      unless border_style_node.xpath('xmlns:bottom')[0].attribute('style').nil?
        cell_borders.bottom = Border.new
        cell_borders.bottom.style = border_style_node.xpath('xmlns:bottom')[0].attribute('style').value
        cell_borders.bottom.color = Color.parse_color_tag(border_style_node.xpath('xmlns:bottom/xmlns:color')[0])
      end
      unless border_style_node.xpath('xmlns:top')[0].attribute('style').nil?
        cell_borders.top = Border.new
        cell_borders.top.style = border_style_node.xpath('xmlns:top')[0].attribute('style').value
        cell_borders.top.color = Color.parse_color_tag(border_style_node.xpath('xmlns:top/xmlns:color')[0])
      end
      unless border_style_node.xpath('xmlns:right')[0].attribute('style').nil?
        cell_borders.right = Border.new
        cell_borders.right.style = border_style_node.xpath('xmlns:right')[0].attribute('style').value
        cell_borders.right.color = Color.parse_color_tag(border_style_node.xpath('xmlns:right/xmlns:color')[0])
      end
      unless border_style_node.xpath('xmlns:left')[0].attribute('style').nil?
        cell_borders.left = Border.new
        cell_borders.left.style = border_style_node.xpath('xmlns:left')[0].attribute('style').value
        cell_borders.left.color = Color.parse_color_tag(border_style_node.xpath('xmlns:left/xmlns:color')[0])
      end
      cell_borders
    end
  end
end
