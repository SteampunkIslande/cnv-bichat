template = """
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CNV Report</title>
    <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
</head>

<body>
    <div class="container mt-5">
        <h1>CNV Report</h1>
        <ul class="nav nav-tabs" id="sampleTabs" role="tablist">
            {% for sample_name, sample_data in samples.items() %}
            <li class="nav-item">
                <a class="nav-link {% if loop.first %}active{% endif %}" id="A{{ sample_name }}-tab" data-toggle="tab"
                    href="#A{{ sample_name }}" role="tab" aria-controls="A{{ sample_name }}"
                    aria-selected="{% if loop.first %}true{% else %}false{% endif %}">{{ sample_name }}</a>
            </li>
            {% endfor %}
        </ul>
        <div class="tab-content" id="sampleTabsContent">
            {% for sample_name, sample_data in samples.items() %}
            <div class="tab-pane fade {% if loop.first %}show active{% endif %}" id="A{{ sample_name }}" role="tabpanel"
                aria-labelledby="A{{ sample_name }}-tab">
                <h3>{{ sample_name }}</h3>
                <h4>Deletions</h4>
                <ul>
                    {% if not sample_data.deletion %}
                    <p>No deletion</p>
                    {% else %}
                    {% for deletion in sample_data.deletion %}
                    <li>{{ deletion }}</li>
                    {% endfor %}
                    {% endif %}
                </ul>
                <h4>Duplications</h4>
                <ul>
                    {% if not sample_data.duplication %}
                    <p>No duplication</p>
                    {% else %}
                    {% for duplication in sample_data.duplication %}
                    <li>{{ duplication }}</li>
                    {% endfor %}
                    {% endif %}
                </ul>
                <h4>Graphe correspondant</h4>
                <img src='data:image/png;base64,{{ sample_data.graph }}' alt='graph'>
            </div>
            {% endfor %}
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
</body>

</html>
</li>
"""
