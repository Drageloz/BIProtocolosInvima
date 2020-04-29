from django.shortcuts import render
from plotly.offline import plot
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# this part is for ploting scatter graph usin plotly

def tiempoProtocolos(request):
    def scatter():
        fig = make_subplots(rows=1, cols=2)
        x1 = [1,2,3,4]
        y1 = [30, 35, 25, 45]

        trace = go.Scatter(
            x=x1,
            y = y1
        )
        layout = dict(
            title='Simple Graph',
            xaxis=dict(range=[min(x1), max(x1)]),
            yaxis = dict(range=[min(y1), max(y1)])
        )

        fig = go.Figure(data=[trace], layout=layout)
        plot_div = plot(fig, output_type='div', include_plotlyjs=False)
        return plot_div

    def barChar():
        # Create random data with numpy
        import numpy as np
        np.random.seed(1)

        N = 100
        random_x = np.linspace(0, 1, N)
        random_y0 = np.random.randn(N) + 5
        random_y1 = np.random.randn(N)
        random_y2 = np.random.randn(N) - 5

        # Create traces
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=random_x, y=random_y0,
                            mode='lines',
                            name='lines'))
        fig.add_trace(go.Scatter(x=random_x, y=random_y1,
                            mode='lines+markers',
                            name='lines+markers'))
        fig.add_trace(go.Scatter(x=random_x, y=random_y2,
                            mode='markers', name='markers'))
        
        plot_div = plot(fig, output_type='div', include_plotlyjs=False)
        return plot_div

    context ={
        'plot1': scatter(),
        'plot2': barChar(),
    }

    return render(request, 'tiempoProtocolos/welcome.html', context)
