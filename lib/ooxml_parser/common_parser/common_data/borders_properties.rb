# Border Properties Data
module OoxmlParser
  class BordersProperties < OOXMLDocumentObject
    attr_accessor :color, :space, :sz, :val, :space, :shadow, :frame, :side

    def initialize(color = :auto, sz = 0, val = :none, space = 0, parent: nil)
      @color = color
      @sz = sz
      @val = val
      @space = space
      @parent = parent
    end

    alias size sz

    def nil?
      @sz.zero? && val == :none
    end

    def to_s
      return '' if nil?
      "borders color: #{@color}, size: #{@sz}, space: #{@space}, value: #{@val}"
    end

    def copy
      BordersProperties.new(@color, @sz, @val, @space)
    end

    def visible?
      return false if nil?
      !(val == 'none')
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with BordersProperties
    # @return [BordersProperties] value of BordersProperties
    def parse(node)
      return nil if node.attribute('val').value == 'nil'
      node.attributes.each do |key, value|
        case key
        when 'val'
          @val = value.value.to_sym
        when 'sz'
          @sz = OoxmlSize.new(value.value.to_f, :one_eighth_point)
        when 'space'
          @space = OoxmlSize.new(value.value.to_f, :point)
        when 'color'
          @color = value.value.to_s
          @color = Color.from_int16(@color) if @color != 'auto'
        when 'shadow'
          @shadow = value.value
        end
      end
      self
    end
  end
end
