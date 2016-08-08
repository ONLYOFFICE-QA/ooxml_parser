module OoxmlParser
  class Configuration
    attr_accessor :units
    # @return [Integer] accuracy of digits in fraction part
    attr_accessor :accuracy

    def initialize
      @units = :centimeters
      @accuracy = 2
    end

    # @return [Float] by which divide units values
    def units_delimiter
      return (20 * 635.0) if @units == :points
      return (566.929 * 635.0) if @units == :centimeters
      return 1 if @units == :emu
      return 1440 if @units == :inches
      return 635.0 if @units == :dxa
      warn "Cannot recognize #{@units} unit. Will use dxa by default"
      1
    end
  end
end
