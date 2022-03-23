Imports YahooAccessData.MathPlus.Filter

Public Interface IStochastic1
  Inherits IStochastic
  ReadOnly Property AsIStochasticPriceGain As IStochasticPriceGain
  ReadOnly Property AsIStochastic1 As IStochastic1

  ReadOnly Property AsIStochastic As IStochastic

End Interface