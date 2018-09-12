# Border Properties Data
module OoxmlParser
  class BordersProperties < OOXMLDocumentObject
    attr_accessor :color, :space, :val, :shadow, :frame, :side
    # @return [OoxmlSize] size of border
    attr_reader :size

    def initialize(color = :auto, size = 0, val = :none, space = 0, parent: nil)
      @color = color
      @size = size
      @val = val
      @space = space
      @parent = parent
    end

    # @return [OoxmlSize] alias for sz
    def sz
      size
    end

    extend Gem::Deprecate
    deprecate :sz, 'size', 2020, 1

    def nil?
      size.zero? && val == :none
    end

    def to_s
      return '' if nil?

      "borders color: #{@color}, size: #{size}, space: #{@space}, value: #{@val}"
    end

    def copy
      BordersProperties.new(@color, size, @val, @space)
    end

    def visible?
      return false if nil?

      val != 'none'
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
          @size = OoxmlSize.new(value.value.to_f, :one_eighth_point)
        when 'space'
          @space = OoxmlSize.new(value.value.to_f, :point)
        when 'color'
          @color = value.value.to_s
          @color = Color.new(parent: self).parse_hex_string(@color) if @color != 'auto'
        when 'shadow'
          @shadow = value.value
        end
      end
      self
    end
  end
end
