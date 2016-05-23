module OoxmlParser
  class LinearGradient
    attr_accessor :angle, :scaled

    def initialize(angle = nil, scaled = nil)
      @angle = angle
      @scaled = scaled
    end

    # Parse LinearGradient object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LinearGradient] result of parsing
    def self.parse(node)
      gradient = LinearGradient.new
      gradient.angle = node.attribute('ang').value.to_f / 100_000.0 if node.attribute('ang')
      gradient.scaled = node.attribute('scaled').value.to_i if node.attribute('scaled')
      gradient
    end
  end
end
